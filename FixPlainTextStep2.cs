using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using Interop.MIMETranslator;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Exchange.WebServices.Autodiscover;
using System.Xml;
using System.Runtime.InteropServices;
using System.Net;

namespace EWSTestApp
{
    class FixPlainTextStep2
    {
        private MIMEToMSG Translator = null;
        private String EmailsFolder = String.Empty;
        private List<ExtendedPropertyDefinition> ExtendedProps = new List<ExtendedPropertyDefinition>();
        Dictionary<String, Dictionary<String, DocumentInfo>> DocumentInfos = new Dictionary<String, Dictionary<String, DocumentInfo>>();
        private Dictionary<String, List<String>> m_oMarkedFolders = null;
        private PropertySet PropsToFetch = null;
        private bool Exchange2007 = false;
        private String ExchangeServerName = String.Empty;
        private ExchangeVersion ExchangeServerVer = ExchangeVersion.Exchange2010_SP2;
        StreamWriter UnresolvedDocuments = new StreamWriter("UnresolvedDocuments.txt", false);
        StreamWriter FailedDocuments = new StreamWriter("FailedDocuments.txt", false);
        StreamWriter Log = new StreamWriter("EWSTestAppLog.txt", true);

        public void Execute(string[] args)
        {
            try
            {
                do
                {
                    Log.AutoFlush = true;

                    if (args.Length < 4)
                    {
                        Log.WriteLine("Invalid parameters");
                        Console.WriteLine("Syntax: <Command> <ImpersonatorSMTP> <password> <Step1 report from MsgFix> [SRV:exchange server name] [EMP:<Exported EM_PROJECTS table>] [VER:<Exchange Server Version>]");
                        Console.WriteLine("Example: FIXPLAINTEXT-STEP2 ImpersonatorSMTPAddress@dev.local password c:\\Step1Report.txt SRV:xchange.dev.local EMP:c:\\EM_Projects.csv Exchange2007_SP1");
                        break;
                    }

                    int UsersCount = 0;
                    int UsersPassed = 0;
                    int UsersFailed = 0;

                    String sImpersonatorSMTP = args[1];
                    String sImpersonatorPwd = args[2];
                    String sReportFilePath = args[3];

                    String sEMProjects = String.Empty;
                    String sExchangeVer = String.Empty;
                    String sTemp = String.Empty;

                    bool bInternalError = false;
                    for (int nIter = 4; nIter < args.Length; nIter++)
                    {
                        sTemp = args[nIter].Trim();
                        if (3 >= sTemp.Length)
                        {
                            bInternalError = true;
                            Console.WriteLine("Invalid param: " + sTemp);
                            break;
                        }

                        if (0 == String.Compare("SRV:", sTemp.Substring(0, 4), true))
                        {
                            ExchangeServerName = sTemp.Substring(4).Trim();
                            if (String.Empty == ExchangeServerName)
                            {
                                bInternalError = true;
                                Console.WriteLine("Invalid Param: Server Name");
                                break;
                            }
                        }
                        else if (0 == String.Compare("EMP:", sTemp.Substring(0, 4), true))
                        {
                            sEMProjects = sTemp.Substring(4).Trim();
                            if (String.Empty == sEMProjects)
                            {
                                bInternalError = true;
                                Console.WriteLine("Invalid Param: EM Projects File");
                                break;
                            }
                        }
                        else if (0 == String.Compare("VER:", sTemp.Substring(0, 4), true))
                        {
                            sExchangeVer = sTemp.Substring(4).Trim();
                            if (String.Empty == sExchangeVer)
                            {
                                bInternalError = true;
                                Console.WriteLine("Invalid Param: Exchange Server Version");
                                break;
                            }
                        }
                        else
                        {
                            bInternalError = true;
                            Console.WriteLine("Invalid Param: " + sTemp);
                            break;
                        }
                    }

                    if (bInternalError)
                    { break; }

                    this.ExchangeServerVer = StringToExchangeVersion(sExchangeVer);
                    Exchange2007 = (ExchangeVersion.Exchange2007_SP1 == this.ExchangeServerVer);

                    if (Exchange2007)
                    {
                        if (String.Empty == sEMProjects)
                        {
                            Console.Write("For Exchange 2007 environments, please provide the EM Projects table entries");
                            break;
                        }

                        m_oMarkedFolders = new Dictionary<String, List<String>>();
                        if (!LoadEMProjects(sEMProjects))
                        {
                            Console.WriteLine(String.Format("Could not load {0} or the file is empty", sEMProjects));
                            break;
                        }
                    }

                    Log.WriteLine(String.Format("ExchangeServerName: {0}", ExchangeServerName));
                    Log.WriteLine(String.Format("Exchange2007: {0}", Exchange2007.ToString()));
                    Log.WriteLine(String.Format("sEMProjects: {0}", sEMProjects));
                    
                    if (!File.Exists(sReportFilePath))
                    {
                        Console.WriteLine(String.Format("Invalid file path. ({0})", sReportFilePath));
                        break;
                    }

                    Log.WriteLine(String.Format("Loading file {0}", sReportFilePath));
                    if (!LoadDocumentInfos(sReportFilePath))
                    {
                        Console.WriteLine(String.Format("Failed to load file {0}", sReportFilePath));
                        break;
                    }

                    UsersCount = DocumentInfos.Keys.Count;
                    if (0 >= UsersCount)
                    {
                        Console.WriteLine("Could not find any users to process");
                        break;
                    }

                    Log.WriteLine("Number of users to process = " + UsersCount);

                    // Create Emails folder
                    String CurrentFolder = Directory.GetCurrentDirectory();
                    EmailsFolder = Path.Combine(CurrentFolder, "E-Mails");

                    if (!Directory.Exists(EmailsFolder))
                    {
                        Directory.CreateDirectory(EmailsFolder);
                        if (!Directory.Exists(EmailsFolder))
                        {
                            Console.WriteLine("Failed to create E-Mails folder under " + CurrentFolder);
                            break;
                        }
                    }

                    Log.WriteLine("Creating Mime Translator");
                    Translator = new MIMEToMSG();
                    if (null == Translator)
                    {
                        Console.WriteLine("Failed to create MIME Translator");
                        break;
                    }

                    Log.WriteLine("Loading Extended Properties");
                    if (!LoadExtendedProperties(ref ExtendedProps))
                    {
                        Console.WriteLine("Could not load GetItemRequest.xml. Please ensure the file is in the current directory.");
                        break;
                    }

                    Log.WriteLine("Preparing list of properties to fetch");
                    PropsToFetch = new PropertySet();
                    PropsToFetch.Add(ItemSchema.MimeContent);
                    PropsToFetch.Add(ItemSchema.Subject);
                    foreach (ExtendedPropertyDefinition ExProp in ExtendedProps)
                    {
                        PropsToFetch.Add(ExProp);
                    }

                    UnresolvedDocuments.AutoFlush = true;
                    FailedDocuments.AutoFlush = true;                    

                    UnresolvedDocuments.WriteLine("Following are the document numbers that resolve to multiple emails in the mailbox");
                    UnresolvedDocuments.WriteLine("--------------------------------------------------------------------------------------");

                    FailedDocuments.WriteLine("Following are the document numbers that could not be processed");
                    FailedDocuments.WriteLine("--------------------------------------------------------------------------------------");

                    // Process users                
                    Log.WriteLine("Processing Users");
                    foreach (KeyValuePair<String, Dictionary<String, DocumentInfo>> Entry in DocumentInfos)
                    {
                        String UserSmtp = Entry.Key;
                        Dictionary<String, DocumentInfo> UserDocuments = Entry.Value;

                        Log.WriteLine("\n");
                        Log.WriteLine(String.Format("Processing user {0}", UserSmtp));
                        if (!ProcessOneUser(sImpersonatorSMTP, sImpersonatorPwd, UserSmtp, ref UserDocuments))
                        {
                            UsersFailed++;
                            continue;
                        }

                        UsersPassed++;
                    }

                    Console.WriteLine(String.Format("\n\nTotal Users: {0}, Succeeded: {1}, Failed: {2}\n\n", UsersCount, UsersPassed, UsersFailed));
                    Log.WriteLine(String.Format("\n\nTotal Users: {0}, Succeeded: {1}, Failed: {2}\n\n", UsersCount, UsersPassed, UsersFailed));
                }
                while (false);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
            }
            finally
            {
                if (null != Translator) { Marshal.ReleaseComObject(Translator); Translator = null; }
                UnresolvedDocuments.Flush(); UnresolvedDocuments.Close();
                FailedDocuments.Flush(); FailedDocuments.Close();
                Log.Flush(); Log.Close();
            }            
        }

        private bool LoadDocumentInfos(String sDocInfoFilePath)
        {
            Log.WriteLine("> LoadDocumentInfos");

            bool bRet = false;
            try
            {
                do 
                {
                    String[] sReportLines = File.ReadAllLines(sDocInfoFilePath);
                    if (0 == sReportLines.Length)
                    {
                        Console.WriteLine(sDocInfoFilePath + " is empty");
                        break;
                    }

                    // Load report file
                    bool bInternalError = false;
                    foreach (String sReportLine in sReportLines)
                    {
                        sReportLine.Trim();

                        if ((0 == sReportLine.Length) || (';' == sReportLine[0]))
                        { continue; }

                        String[] sTokens = sReportLine.Split("|".ToCharArray());

                        // 2395|SSAHOO@BLRDEV.NET|AF3B3317-FEFF-4C82-8AFA-ED0F97E48785|\t\t- Contains Garbage
                        if (3 > sTokens.Length)
                        { continue; }

                        String DocNum = sTokens[0];
                        String UserSmtp = sTokens[1];
                        String AutnGuid = sTokens[2];

                        if ((String.Empty == DocNum) || (String.Empty == UserSmtp) || (String.Empty == AutnGuid))
                        { continue; }

                        DocumentInfo docInfo = new DocumentInfo(DocNum, UserSmtp, AutnGuid);

                        if (!DocumentInfos.ContainsKey(UserSmtp))
                        {
                            // User not already added
                            Dictionary<string, DocumentInfo> newUserInfo = new Dictionary<string, DocumentInfo>();
                            newUserInfo.Add(DocNum, docInfo);
                            DocumentInfos.Add(UserSmtp, newUserInfo);
                        }
                        else
                        {
                            Dictionary<String, DocumentInfo> UserDocuments = DocumentInfos[UserSmtp];
                            if (!UserDocuments.ContainsKey(DocNum))
                            {
                                UserDocuments.Add(DocNum, docInfo);
                            }
                            else 
                            {
                                Log.WriteLine(String.Format("Found duplicate DocNum {0} for user {1}. Skipping", DocNum, UserSmtp));
                                Console.WriteLine(String.Format("Found duplicate DocNum {0} for user {1}. Skipping.", DocNum, UserSmtp));
                            }
                        }
                    }
                    bRet = !bInternalError;
                }
                while (false);                
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                bRet = false;
            }

            Log.WriteLine("< LoadDocumentInfos");
            return bRet;
        }

        private bool LoadEMProjects(String CSVFilePath)
        {
            bool bRet = false;
            do
            {
                if (!File.Exists(CSVFilePath))
                { break; }

                String[] EMProjectEntries = File.ReadAllLines(CSVFilePath);
                if (0 == EMProjectEntries.Length)
                { break; }

                Int32 sLineNum = 0;
                foreach (String sEntry in EMProjectEntries)
                {
                    sEntry.Trim();
                    if (String.IsNullOrEmpty(sEntry))
                    { continue; }

                    String[] Tokens = sEntry.Split(",".ToCharArray());
                    if (2 > Tokens.Length)
                    { throw new Exception(String.Format("Invalid entry in {0} at line {1}", CSVFilePath, sLineNum)); }

                    String sUserSMTP = Tokens[0];
                    String sFolderEntry = Tokens[1];

                    int nIndex1 = sUserSMTP.IndexOf('@');
                    if (1 > nIndex1)
                    { continue; }

                    int nIndex2 = sUserSMTP.Substring(nIndex1 + 1).IndexOf('.');
                    if (1 > nIndex2)
                    { continue; }

                    if (!m_oMarkedFolders.ContainsKey(sUserSMTP))
                    {
                        List<String> UserMarkedFolder = new List<String>();
                        UserMarkedFolder.Add(sFolderEntry);
                        m_oMarkedFolders.Add(sUserSMTP, UserMarkedFolder);
                    }
                    else 
                    {
                        List<String> UserMarkedFolders = m_oMarkedFolders[sUserSMTP];
                        try 
                        { 
                            UserMarkedFolders.Add(sFolderEntry); 
                        }
                        catch 
                        { 
                            Log.WriteLine(String.Format("Found duplicate folder entry id {0} for user {1}", sFolderEntry, sUserSMTP)); 
                        }
                    }
                }
                bRet = true;
            }
            while (false);
            return bRet;
        }

        private bool PrepareFolderListToSearch(ref ExchangeService service, String sUserSMTP, ref List<FolderId> FolderIds)
        {
            bool bRet = false;

            do
            {
                // If a list of folder IDs is provided, then limit the search
                // to only those folders along with the default folders.
                // Otherwise, search AllItems folder.

                if (!Exchange2007)
                {
                    Folder AllItems = null;
                    bRet = GetAllItemsFolder(ref service, ref AllItems);
                    if (null == AllItems)
                    {
                        bRet = false;
                        break;
                    }
                    
                    FolderIds.Add(AllItems.Id);
                    bRet = true;
                    break;
                }

                FolderId InboxId = null;
                FolderId SentItemsId = null;
                FolderId DeletedItemsId = null;

                InboxId = GetWellKnownFolderId(ref service, WellKnownFolderName.Inbox);
                if (null == InboxId)
                {
                    Console.WriteLine("Could not get EWS Id for Inbox");
                    break;
                }

                SentItemsId = GetWellKnownFolderId(ref service, WellKnownFolderName.SentItems);
                if (null == SentItemsId)
                {
                    Console.WriteLine("Could not get EWS Id for Sent Items");
                    break;
                }

                DeletedItemsId = GetWellKnownFolderId(ref service, WellKnownFolderName.DeletedItems);
                if (null == DeletedItemsId)
                {
                    Console.WriteLine("Could not get EWS Id for Deleted Items");
                    break;
                }

                if (null == FolderIds)
                {
                    FolderIds = new List<FolderId>();
                }

                // Maintain the search order: firs the default folders. Then marked folders.
                FolderIds.Add(InboxId);
                FolderIds.Add(DeletedItemsId);
                FolderIds.Add(SentItemsId);

                if (!m_oMarkedFolders.ContainsKey(sUserSMTP))
                {
                    // May be the user does not have any marked folders.
                    // We will search only the default folders
                    bRet = true;
                    break;
                }

                List<String> UserMarkedFolders = m_oMarkedFolders[sUserSMTP];
                foreach (String markedFolderEntryId in UserMarkedFolders)
                {
                    FolderId markedFolderId = null;
                    markedFolderId = GetMarkedFolderId(ref service, sUserSMTP, markedFolderEntryId);
                    if (null == FolderIds)
                    {
                        Log.WriteLine("Could not find folder for entry id: " + markedFolderEntryId);
                        continue;
                    }

                    FolderIds.Add(markedFolderId);                    
                }
                bRet = true;
            }
            while (false);
            return bRet;
        }

        private bool ConnectToExchangeServer(String sImpersonatorSMTP, String sImpersonatorPwd, String sUserSMTP, ref ExchangeService service)
        {
            bool bRet = false;
            Log.WriteLine("> ConnectToExchangeServer");

            try
            {
                Log.WriteLine("Connecting Exchange server");
                service = new ExchangeService(this.ExchangeServerVer);
                
                service.Credentials = new WebCredentials(sImpersonatorSMTP, sImpersonatorPwd);
                service.TraceListener = new TraceListener();
                service.TraceFlags = TraceFlags.All;
                service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, sUserSMTP);
                service.TraceEnabled = true;
                ServicePointManager.ServerCertificateValidationCallback = Program.CertificateValidationCallback;

                if (String.Empty != ExchangeServerName)
                {
                    string exchangeUrl;
                    exchangeUrl = "https://";
                    exchangeUrl += ExchangeServerName;
                    exchangeUrl += "/EWS/Exchange.asmx";
                    service.Url = new Uri(exchangeUrl);
                }
                else
                {
                    Log.WriteLine("Calling AutodiscoverUrl");
                    service.AutodiscoverUrl(sUserSMTP, Program.RedirectionUrlValidationCallback);
                }

                PropertySet p = new PropertySet(BasePropertySet.IdOnly);
                Folder ewsFolder = Folder.Bind(service, WellKnownFolderName.Inbox, p);
                bRet = (null != ewsFolder);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to connect to the exchange server. " + ex.Message);
                bRet = false;
            }

            Log.WriteLine("< ConnectToExchangeServer");
            return bRet;
        }

        private bool ProcessOneUser(String sImpersonatorSMTP, String sImpersonatorPwd, String UserSmtp, ref Dictionary<String, DocumentInfo> UserDocuments)
        {
            bool bRet = false;

            Console.WriteLine("Processing user " + UserSmtp);
            UnresolvedDocuments.WriteLine("User: " + UserSmtp);
            FailedDocuments.WriteLine("User: " + UserSmtp);

            do
            {
                int TotalDocs = UserDocuments.Count;
                int PassedDocs = 0;
                int FailedDocs = 0;

                Log.WriteLine("Connecting to exchange server");
                ExchangeService service = null;
                bool bConnected = ConnectToExchangeServer(sImpersonatorSMTP, sImpersonatorPwd, UserSmtp, ref service);
                
                if (!bConnected || (null == service) || (null == service.Url))
                {
                    Console.WriteLine("Failed to connect to exchange server for user " + UserSmtp);
                    break;
                }

                String AutoDiscoverURL = service.Url.ToString();
                AutoDiscoverURL = AutoDiscoverURL.Trim();

                if (String.Empty == AutoDiscoverURL)
                {
                    Console.WriteLine("Failed to get exchange server for user " + UserSmtp);
                    break;
                }

                Log.WriteLine("Preparing Folder List to search");
                List<FolderId> FoldersToSearch = new List<FolderId>();
                if (!PrepareFolderListToSearch(ref service, UserSmtp, ref FoldersToSearch))
                {
                    Console.WriteLine("Failed to prepare folder list for searching");
                    break;
                }

                Log.WriteLine("Processing documents for user");
                bool InternalError = false;
                foreach (KeyValuePair<String, DocumentInfo> UserDocument in UserDocuments)
                {
                    DocumentInfo docInfo = UserDocument.Value;
                    bool bProcessedDocument = false;
                    bool bDupsMismatch = false;

                    foreach (FolderId folderToSearch in FoldersToSearch)
                    {
                        try
                        {
                            FolderId SearchFolderId = folderToSearch;
                            bProcessedDocument = ProcessOneDocument(ref service, ref SearchFolderId, ref docInfo, ref bDupsMismatch);
                            if (bProcessedDocument || bDupsMismatch)
                            { break; }
                        }
                        catch (Exception ex)
                        {
                            InternalError = true; ;
                            Log.WriteLine("\nCannot continue to process this user.\nError:" + ex.Message + ex.StackTrace);
                            Console.WriteLine("\nCannot continue to process this user.\nError:" + ex.Message);
                            Console.WriteLine();
                            break;
                        }
                    }

                    if (bDupsMismatch)
                    {
                        Log.WriteLine("Could not resolve document " + docInfo.DocNum);
                        UnresolvedDocuments.WriteLine(docInfo.DocNum);
                        FailedDocs++;
                        continue;
                    }

                    if (!bProcessedDocument)
                    {
                        Log.WriteLine("Failed to process document " + docInfo.DocNum);
                        FailedDocuments.WriteLine(docInfo.DocNum);
                        FailedDocs++;
                        continue;
                    }

                    Log.WriteLine("Processed document " + docInfo.DocNum);
                    Console.WriteLine("Processed document " + docInfo.DocNum);
                    PassedDocs++;
                }

                Log.WriteLine(String.Format("\nTotal Documents {0}, Succeeded {1}, Failed {2}", TotalDocs, PassedDocs, FailedDocs));
                Console.WriteLine(String.Format("\nTotal Documents {0}, Succeeded {1}, Failed {2}", TotalDocs, PassedDocs, FailedDocs));
                bRet = !InternalError;

            }
            while (false);
            Log.WriteLine("< ProcessOneUser");
            return bRet;
        }

        private FolderId GetWellKnownFolderId(ref ExchangeService service, WellKnownFolderName folderName)
        {
            Log.WriteLine("> GetWellKnownFolderId");

            FolderId folderId = null;

            FolderView findFolderView = new FolderView(1000);
            findFolderView.Traversal = FolderTraversal.Shallow;

            Folder WellKnownFolder = Folder.Bind(service, folderName);
            //FindFoldersResults findFolderRes = service.FindFolders(folderName, findFolderView);
            if ((null != WellKnownFolder) /*&& (1 == findFolderRes.Folders.Count)*/)
            {
                //folderId = findFolderRes.Folders[0].Id;
                folderId = WellKnownFolder.Id;
            }

            Log.WriteLine("< GetWellKnownFolderId");
            return folderId;
        }

        private FolderId GetMarkedFolderId(ref ExchangeService service, String UserSmtp, String sFolderEntryID)
        {
            FolderId folderId = null;
            Log.WriteLine("> GetMarkedFolderId");

            do
            {
                String FolderEWSId = Program.GetConvertedEWSID(service, sFolderEntryID, UserSmtp);
                if (String.Empty == FolderEWSId)
                {
                    Console.WriteLine(String.Format("Failed to get folder for id {0}"), sFolderEntryID);
                    break;
                }

                PropertySet folderprops = new PropertySet(BasePropertySet.IdOnly);
                folderprops.Add(FolderSchema.DisplayName);
                Folder MarkedFolder = Folder.Bind(service, new FolderId(FolderEWSId));
                if (null == MarkedFolder)
                { break; }

                folderId = MarkedFolder.Id;                
            }
            while (false);

            Log.WriteLine("< GetMarkedFolderId");
            return folderId;
        }
        
        private bool ProcessOneDocument(ref ExchangeService service, ref FolderId AllItemsFolder, ref DocumentInfo docInfo, ref bool DupsMismatch)
        {
            bool bRet = false;
            DupsMismatch = false;

            String EmailPath = String.Empty;
            String MsgPath = String.Empty;
            String MetadataPath = String.Empty;
            String ExPropsPath = String.Empty;
            String FirstFilePath = String.Empty;

            Log.WriteLine(">ProcessOneDocument: Processing document: " + docInfo.DocNum);

            try
            {
                do
                {
                    String EmailGuid = docInfo.AutnGuid;
                    Item OutlookItem = null;

                    ExtendedPropertyDefinition autnguid = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders, "x-autn-guid", MapiPropertyType.String);
                    ItemView view = new ItemView(10);
                    view.PropertySet = new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.Subject);
                    view.Offset = 0;
                    view.OffsetBasePoint = OffsetBasePoint.Beginning;
                    SearchFilter.SearchFilterCollection filter = new SearchFilter.SearchFilterCollection();
                    filter.Add(new SearchFilter.IsEqualTo(autnguid, EmailGuid));

                    Log.WriteLine("ProcessOneDocument: Calling FindItems");
                    FindItemsResults<Item> findResults = null;
                    try
                    {
                        findResults = service.FindItems(AllItemsFolder, filter, view);
                    }
                    catch (Exception ex)
                    {
                        Log.WriteLine("ProcessOneDocument: Exception (FindItems). " + ex.Message + ex.StackTrace);
                        bRet = false; 
                        break;
                    }

                    if (null == findResults)
                    {
                        Log.WriteLine("ProcessOneDocument: finResults is null");
                        bRet = false; 
                        break;
                    }

                    if (0 >= findResults.Items.Count)
                    {
                        Log.WriteLine("ProcessOneDocument: Could not find email with guid " + EmailGuid);
                        bRet = false; 
                        break;
                    }

                    Log.WriteLine("ProcessOneDocument: Calling BindToItems");
                    ServiceResponseCollection<GetItemResponse> bindResults = null;
                    try
                    {
                        bindResults = service.BindToItems(findResults.Select(r => r.Id), PropsToFetch);
                    }
                    catch (Exception ex)
                    {
                        Log.WriteLine("ProcessOneDocument: Exception (BindToItems). " + ex.Message + ex.StackTrace);
                        bRet = false; 
                        break; 
                    }

                    if ((null == bindResults) || (0 >= bindResults.Count))
                    {
                        Log.WriteLine("ProcessOneDocument: Failed to bind to the email");
                        bRet = false; 
                        break; 
                    }

                    Log.WriteLine("ProcessOneDocument: BindToItems succeeded");
                    if (1 < bindResults.Count)
                    {
                        Log.WriteLine("ProcessOneDocument: More than one emails match the guid [" + EmailGuid + "]");
                    }

                    for (int nIter = 0; nIter < bindResults.Count; nIter++)
                    {
                        OutlookItem = bindResults[nIter].Item;
                        String FilePath = String.Empty;
                        if (0 == nIter)
                        {
                            FilePath = Path.Combine(EmailsFolder, String.Format("{0}_{1}", docInfo.DocNum, docInfo.AutnGuid));
                        }
                        else
                        {
                            FilePath = Path.Combine(EmailsFolder, String.Format("{0}_{1}_{2}", docInfo.DocNum, docInfo.AutnGuid, nIter));
                        }

                        EmailPath = FilePath + ".eml";

                        Log.WriteLine("ProcessOneDocument: Delete if exists: " + EmailPath);
                        if (File.Exists(EmailPath))
                        {
                            File.Delete(EmailPath);
                            if (File.Exists(EmailPath))
                            {
                                Log.WriteLine(String.Format("ProcessOneDocument: File {0} already exists", EmailPath));
                                break;
                            }
                        }

                        Log.WriteLine("ProcessOneDocument: Writing MIMEContent to file. nIter = " + nIter);
                        File.WriteAllText(EmailPath, OutlookItem.MimeContent.ToString(), Encoding.UTF8);

                        if (0 == nIter)
                        {
                            MsgPath = FilePath + ".msg";
                            MetadataPath = FilePath + ".xml";
                            ExPropsPath = FilePath + "_XP.xml";

                            if (File.Exists(MsgPath))
                            {
                                File.Delete(MsgPath);
                                if (File.Exists(MsgPath))
                                {
                                    Log.WriteLine(String.Format("ProcessOneDocument: File {0} already exists", MsgPath));
                                    break;
                                }
                            }

                            if (File.Exists(MetadataPath))
                            {
                                File.Delete(MetadataPath);
                                if (File.Exists(MetadataPath))
                                {
                                    Log.WriteLine(String.Format("ProcessOneDocument: File {0} already exists", MetadataPath));
                                    break;
                                }
                            }

                            if (File.Exists(ExPropsPath))
                            {
                                File.Delete(ExPropsPath);
                                if (File.Exists(ExPropsPath))
                                {
                                    Log.WriteLine(String.Format("ProcessOneDocument: File {0} already exists", ExPropsPath));
                                    break;
                                }
                            }

                            FirstFilePath = EmailPath;

                            Log.WriteLine("ProcessOneDocument: Generating Extended Properties XML");
                            GenerateExPropsXML(OutlookItem.ExtendedProperties, ExPropsPath);

                            Log.WriteLine("ProcessOneDocument: Translating file " + EmailPath);
                            Translator.Translate(EmailPath, MsgPath, MetadataPath, ExPropsPath);
                            bRet = true;
                            continue;
                        }

                        Log.WriteLine(String.Format("ProcessOneDocument: Comparing files {0} and {1}", FirstFilePath, EmailPath));
                        bool FilesAreSame = TheFilesAreSame(FirstFilePath, EmailPath);
                        File.Delete(EmailPath);

                        if (!FilesAreSame)
                        {
                            Log.WriteLine("ProcessOneDocument: Emails with duplicate guids are not identical");
                            File.Delete(MsgPath);
                            DupsMismatch = true;
                            bRet = false;
                            break;
                        }
                        else 
                        {
                            Log.WriteLine("ProcessOneDocument: Emails with duplicate guids are identical");
                        }
                    }

                } while (false);
            }
            catch (Exception ex)
            {
                Log.WriteLine("ProcessOneDocument: Exception: " + ex.Message + ex.StackTrace);
            }
            finally 
            {
                Log.WriteLine("ProcessOneDocument: Deleting temporary files");
                try
                {
                    if (File.Exists(EmailPath)) { File.Delete(EmailPath); }
                    if (File.Exists(MetadataPath)) { File.Delete(MetadataPath); }
                    if (File.Exists(ExPropsPath)) { File.Delete(ExPropsPath); }
                    if (File.Exists(FirstFilePath)) { File.Delete(FirstFilePath); }
                }
                catch (Exception ex)
                {
                    Log.WriteLine("ProcessOneDocument: Exception: " + ex.Message + ex.StackTrace);
                }
            }

            Log.WriteLine("< ProcessOneDocument");
            return bRet;
        }

        private bool GetAllItemsFolder(ref ExchangeService service, ref Folder AllItems)
        {
            bool bRet = false;
            Log.WriteLine("> GetAllItemsFolder");

            ExtendedPropertyDefinition prFolderType = new ExtendedPropertyDefinition(13825, MapiPropertyType.Integer);
            SearchFilter.SearchFilterCollection filterAllItemsFolder = new SearchFilter.SearchFilterCollection(LogicalOperator.And);
            filterAllItemsFolder.Add(new SearchFilter.IsEqualTo(prFolderType, "2"));
            filterAllItemsFolder.Add(new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "AllItems"));
            FolderView viewAllItemsFolder = new FolderView(1000);
            viewAllItemsFolder.Traversal = FolderTraversal.Shallow;

            FindFoldersResults findAllItemsFolder = service.FindFolders(WellKnownFolderName.Root, filterAllItemsFolder, viewAllItemsFolder);
            if (findAllItemsFolder.Folders.Count > 0)
            {
                AllItems = findAllItemsFolder.Folders[0];
                bRet = true;
            }

            Log.WriteLine("< GetAllItemsFolder");
            return bRet;
        }

        private bool GetItemByGuid(ref ExchangeService service, ref FolderId AllItemsFolder, String sEmailGuid, ref Item OutlookItem, ref bool MultiplesFound)
        {
            bool bRet = false;
            OutlookItem = null;
            MultiplesFound = false;

            do
            {
                ExtendedPropertyDefinition autnguid = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders, "x-autn-guid", MapiPropertyType.String);

                ItemView view = new ItemView(10);
                view.PropertySet = new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.Subject);
                view.Offset = 0;
                view.OffsetBasePoint = OffsetBasePoint.Beginning;

                SearchFilter.SearchFilterCollection filter = new SearchFilter.SearchFilterCollection(LogicalOperator.Or);
                filter.Add(new SearchFilter.IsEqualTo(autnguid, sEmailGuid));

                FindItemsResults<Item> findResults = null;
                try
                {
                    findResults = service.FindItems(AllItemsFolder, filter, view);
                }
                catch { break; }

                ServiceResponseCollection<GetItemResponse> bindResults = null;

                try
                {
                    bindResults = service.BindToItems(findResults.Select(r => r.Id), PropsToFetch);
                }
                catch { break; }

                if ((null == bindResults) || (0 >= bindResults.Count))
                { break; }
                else if (1 < bindResults.Count)
                {
                    // Found more than one item with the same guid
                    MultiplesFound = true;
                    throw new Exception("More than one emails match the guid [" + sEmailGuid + "]");
                }

                OutlookItem = bindResults[0].Item;
                bRet = (null != OutlookItem);

            } while (false);
            return bRet;
        }

        private bool LoadExtendedProperties(ref List<ExtendedPropertyDefinition> ExProps)
        {
            bool bRet = false;
            
            do 
            {
                String GetItemRequestXML = Path.Combine(Directory.GetCurrentDirectory(), "GetItemRequest.xml");
                if (!File.Exists(GetItemRequestXML))
                { break; }

                XmlDocument GIRequest = new XmlDocument();
                XmlNamespaceManager NamespaceManager = new XmlNamespaceManager(GIRequest.NameTable);

                NamespaceManager.AddNamespace("soap", "http://schemas.xmlsoap.org/soap/envelope/");
                NamespaceManager.AddNamespace("t", "http://schemas.microsoft.com/exchange/services/2006/types");
                NamespaceManager.AddNamespace("m", "http://schemas.microsoft.com/exchange/services/2006/messages");
                NamespaceManager.AddNamespace("xsi", "http://www.w3.org/2001/XMLSchema-instance");

                GIRequest.Load(GetItemRequestXML);

                XmlNodeList ExtendedProps = GIRequest.SelectNodes("//soap:Envelope/soap:Body/m:GetItem/m:ItemShape/t:AdditionalProperties/t:ExtendedFieldURI", NamespaceManager);

                if (null == ExtendedProps)
                { break; }

                foreach (XmlNode AdditionalProp in ExtendedProps)
                {
                    String PropertyName = "", PropertyType = "", DistinguishedPropertySetId = "";

                    foreach (XmlAttribute Attribute in AdditionalProp.Attributes)
                    {
                        switch (Attribute.Name)
                        {
                            case "PropertyName":
                                PropertyName = Attribute.Value;
                                break;
                            case "PropertyType":
                                PropertyType = Attribute.Value;
                                break;
                            case "PropertySetId":
                            case "DistinguishedPropertySetId":
                                DistinguishedPropertySetId = Attribute.Value;
                                break;
                        }
                    }

                    try 
                    {
                        Guid PropSetGuid = new Guid(DistinguishedPropertySetId);
                        ExProps.Add(new ExtendedPropertyDefinition(PropSetGuid, PropertyName, StringToMapiPropType(PropertyType)));
                    }
                    catch 
                    {
                        ExProps.Add(new ExtendedPropertyDefinition(StringToPropSet(DistinguishedPropertySetId), PropertyName, StringToMapiPropType(PropertyType)));
                    }                    
                }
                bRet = true;
            }
            while (false);

            return bRet;
        }

        private DefaultExtendedPropertySet StringToPropSet(String sPropSet)
        {
            DefaultExtendedPropertySet PropSet = DefaultExtendedPropertySet.PublicStrings;
            switch (sPropSet)
            { 
                case "Common":
                    PropSet = DefaultExtendedPropertySet.Common;
                    break;
                case "PublicStrings":
                    PropSet = DefaultExtendedPropertySet.PublicStrings;
                    break;
                case "InternetHeaders":
                    PropSet = DefaultExtendedPropertySet.InternetHeaders;
                    break;
                default:
                    throw new Exception("Invalid Property Set");
            }
            return PropSet;
        }

        private MapiPropertyType StringToMapiPropType(String sPropType)
        {
            MapiPropertyType PropType = MapiPropertyType.String;
            switch (sPropType)
            {
                case "Binary":
                    PropType = MapiPropertyType.Binary;
                    break;
                case "Boolean":
                    PropType = MapiPropertyType.Boolean;
                    break;
                case "Integer":
                    PropType = MapiPropertyType.Integer;
                    break;
                case "SystemTime":
                    PropType = MapiPropertyType.SystemTime;
                    break;
                case "String":
                    PropType = MapiPropertyType.String;
                    break;
                default:
                    throw new Exception("Invalid Property Type");
            }
            return PropType;
        }

        private bool GenerateExPropsXML(ExtendedPropertyCollection ExProps, String ExPropsPath)
        {
            bool bRet = false;

            XmlDocument ExPropsXML = new XmlDocument();
            ExPropsXML.LoadXml("<?xml version=\"1.0\" encoding=\"UTF-8\" ?><ExtendedProperties></ExtendedProperties>");
            XmlNode ExtendedPropertiesNode = ExPropsXML.SelectSingleNode("//ExtendedProperties");


            foreach (ExtendedProperty ExProp in ExProps)
            {
                String PropName = ExProp.PropertyDefinition.Name;
                String PropType = MAPIPropTypeToString(ExProp.PropertyDefinition.MapiType);
                String PropSet = String.Empty;
                if (null != ExProp.PropertyDefinition.PropertySet)
                {
                    PropSet = EWSPropSetIdToString((DefaultExtendedPropertySet)ExProp.PropertyDefinition.PropertySet);
                }
                else 
                {
                    if (null != ExProp.PropertyDefinition.PropertySetId)
                    {
                        PropSet = ExProp.PropertyDefinition.PropertySetId.ToString();
                    }
                    else 
                    {
                        throw new Exception("Both PropertySet and PropertySetId are null");
                    }
                }
                String Value = ExProp.Value.ToString();

                XmlNode ExtendedPropertyNode = ExPropsXML.CreateElement("ExtendedProperty");
                ExtendedPropertyNode.InnerText = Value;
                
                XmlAttribute PropertyNameAttr = ExPropsXML.CreateAttribute("PropertyName");
                PropertyNameAttr.Value = PropName;

                XmlAttribute PropertyTypeAttr = ExPropsXML.CreateAttribute("PropertyType");
                PropertyTypeAttr.Value = PropType;

                XmlAttribute PropertySetIdAttr = ExPropsXML.CreateAttribute("PropertySetId");
                PropertySetIdAttr.Value = PropSet;

                ExtendedPropertyNode.Attributes.Append(PropertyNameAttr);
                ExtendedPropertyNode.Attributes.Append(PropertyTypeAttr);
                ExtendedPropertyNode.Attributes.Append(PropertySetIdAttr);

                ExtendedPropertiesNode.AppendChild(ExtendedPropertyNode);
            }

            ExPropsXML.Save(ExPropsPath);
            return bRet;    
        }

        private String MAPIPropTypeToString(MapiPropertyType MapiType)
        {
            String type = String.Empty;
            switch (MapiType.ToString().ToUpper())
            { 
                case "BINARY":
                    type = ((Int32)InternalPropType.Binary).ToString();
                    break;
                case "BOOLEAN":
                    type = ((Int32)InternalPropType.Boolean).ToString();
                    break;
                case "INTEGER":
                    type = ((Int32)InternalPropType.Integer).ToString();
                    break;
                case "SHORT":
                    type = ((Int32)InternalPropType.Short).ToString();
                    break;
                case "LONG":
                    type = ((Int32)InternalPropType.Long).ToString();
                    break;
                case "STRING":
                    type = ((Int32)InternalPropType.String).ToString();
                    break;
                case "SYSTEMTIME":
                    type = ((Int32)InternalPropType.SystemTime).ToString();
                    break;
                default:
                    throw new Exception("Unsupported data type for extended property: " + MapiType.ToString());
            }
            return type;
        }

        private String EWSPropSetIdToString(DefaultExtendedPropertySet EWSPropSet)
        {
            String PropSet = String.Empty;

            switch (EWSPropSet)
            { 
                case DefaultExtendedPropertySet.Common:
                    PropSet = ((Int32)InternalPropSet.PSCommon).ToString();
                    break;
                case DefaultExtendedPropertySet.InternetHeaders:
                    PropSet = ((Int32)InternalPropSet.PSInternetHeaders).ToString();
                    break;
                case DefaultExtendedPropertySet.PublicStrings:
                    PropSet = ((Int32)InternalPropSet.PSPublicStrings).ToString();
                    break;
                default:
                    throw new Exception("Unsupported property set for extended property");
            }

            return PropSet;
        }

        private bool TheFilesAreSame(String sFile1, String sFile2)
        {
            bool bFilesAreSame = true;
            try
            {
                using (Stream file1 = new FileStream(sFile1, FileMode.Open))
                {
                    using (Stream file2 = new FileStream(sFile2, FileMode.Open))
                    {
                        const int bufferSize = 2048;
                        byte[] buffer1 = new byte[bufferSize]; //buffer size                        
                        byte[] buffer2 = new byte[bufferSize];

                        Array.Clear(buffer1, 0, bufferSize);
                        Array.Clear(buffer2, 0, bufferSize);

                        while (true)
                        {
                            int count1 = file1.Read(buffer1, 0, bufferSize);
                            int count2 = file2.Read(buffer2, 0, bufferSize);

                            if (count1 != count2)
                            { bFilesAreSame = false; break; }

                            if (count1 == 0)
                            { bFilesAreSame = true; break; }

                            if (!buffer1.Take(count1).SequenceEqual(buffer2.Take(count2)))
                            { bFilesAreSame = false; break; }
                        }
                    }
                }

            }
            catch
            {
                bFilesAreSame = true;
            }
            return bFilesAreSame;
        }

        private ExchangeVersion StringToExchangeVersion(String sVersion)
        {
            ExchangeVersion Version = ExchangeVersion.Exchange2010_SP2;

            switch (sVersion)
            {
                case "Exchange2007_SP1":
                    Version = ExchangeVersion.Exchange2007_SP1;
                    break;
                case "Exchange2010":
                    Version = ExchangeVersion.Exchange2010;
                    break;
                case "Exchange2010_SP1":
                    Version = ExchangeVersion.Exchange2010_SP1;
                    break;
                case "Exchange2010_SP2":
                default:
                    Version = ExchangeVersion.Exchange2010_SP2;
                    break;
            }

            return Version;
        }
    }

    #region Data Type Definitions
    internal class DocumentInfo
    {
        public DocumentInfo(String sDocNum, String sUserSmtp, String sAutnGuid)
        {
            this.DocNum = sDocNum;
            this.UserSmtp = sUserSmtp;
            this.AutnGuid = sAutnGuid;
        }

        public String DocNum = String.Empty;
        public String UserSmtp = String.Empty;
        public String AutnGuid = String.Empty;
    }

    internal enum InternalPropType
    {
        ApplicationTime = 0,
        ApplicationTimeArray = 1,
        Binary = 2,
        BinaryArray = 3,
        Boolean = 4,
        CLSID_ = 5,
        CLSIDArray = 6,
        Currency = 7,
        CurrencyArray = 8,
        Double = 9,
        DoubleArray = 10,
        Error = 11,
        Float = 12,
        FloatArray = 13,
        Integer = 14,
        IntegerArray = 15,
        Long = 16,
        LongArray = 17,
        Null = 18,
        Object = 19,
        ObjectArray = 20,
        Short = 21,
        ShortArray = 22,
        SystemTime = 23,
        SystemTimeArray = 24,
        String = 25,
        StringArray = 26,
        TypeNone

    };

    internal enum InternalPropSet
    {
        PSMeeting = 0,
        PSAppointment = 1,
        PSCommon = 2,
        PSPublicStrings = 3,
        PSAddress = 4,
        PSInternetHeaders = 5,
        PSCalendarAssistant = 6,
        PSUnifiedMessaging = 7,
        PSTask = 8,
        PSMapiTag = 9,	// To get/set MAPI tags
        PSCustom = 10,	// For Custom Property Set IDs of SEV/EAS Stubs. Validate that it is a GUID.
        PSNone

    };
#endregion
}
