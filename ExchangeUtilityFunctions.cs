using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using Com.Interwoven.WorkSite.iManage;

namespace EWSTestApp
{
    class ExchangeUtilityFunctions
    {
        StreamWriter Log = null; //new StreamWriter("EWSTestAppLog.txt", true);
        StreamWriter EwsReportLog = null;// new StreamWriter("WCSUtilReport.csv", true);
        StreamWriter EwsMappedFolderReportLog = null; // new StreamWriter("WCSUtilReport-MappedFolders.csv", true);
        StreamWriter EwsMFReportLog = null; //new StreamWriter("WCSUtilReportMappedFolders.csv", true);

        // Queued emails per user
        private Dictionary<String, List<ExchangeQueuedEmails>> m_oQueuedEmails = null;

        private Dictionary<string, string> m_oDbConns = null;

        //private Dictionary<String, ExchangeQueuedEmails> m_oEmRequestBucket = new Dictionary<String, ExchangeQueuedEmails>();
        //private Dictionary<String, ExchangeQueuedEmails> m_oNoEmRequestBucket = new Dictionary<String, ExchangeQueuedEmails>();

        WorkSiteUtility m_oWorkSession = null;
        public ExchangeUtilityFunctions(ref WorkSiteUtility workUtility)
        {
            m_oWorkSession = workUtility;
        }

        public void Cleanup()
        {
            if (Log != null)
                Log.Close();

            if (EwsReportLog != null)
                EwsReportLog.Close();

            if (EwsMappedFolderReportLog != null)
                EwsMappedFolderReportLog.Close();

            if (EwsMFReportLog != null)
                EwsMFReportLog.Close();
        }

        public void CreateLogFile(int iLogType, string sLogFileName)
        {
            if (iLogType == 1)
                Log = new StreamWriter(sLogFileName, true);

            if (iLogType == 2)
                EwsReportLog = new StreamWriter(sLogFileName, true);

            if (iLogType == 3)
                EwsMappedFolderReportLog = new StreamWriter(sLogFileName, true);

            if (iLogType == 4)
                EwsMFReportLog = new StreamWriter(sLogFileName, true);
        }

        public bool Initialize(string sUserId, string sPassword)
        {
            bool bRet = true;


            if (!File.Exists("DatabaseConfig.txt"))
            {
                Console.WriteLine("DatabaseConfig.txt doesn't exist");
                Log.WriteLine("DatabaseConfig.txt doesn't exist");
                bRet = false;
            }

            System.IO.StreamReader file = new System.IO.StreamReader("DatabaseConfig.txt");
            string line;

            m_oDbConns = new Dictionary<string, string>();


            string sDbEntry = "";
            int icnt = 0;
            string sDbName = "";
            while ((line = file.ReadLine()) != null)
            {
                try
                {
                    if (line.Length > 0)
                    {

                        if (line.ToUpper() == "[DATABASEINFO]")
                        {
                            bRet = false;
                            if (icnt > 0)
                            {
                                break;
                            }

                            sDbEntry = "";
                            continue;
                        }

                       

                        string[] fileData = line.Split('=');
                        string sKey = fileData[0];
                        string sValue = fileData[1];

                        if (fileData.Count() > 1)
                        {

                            if (sKey.ToUpper() == "DATASOURCE")
                            {
                                sValue = fileData[1];
                                sDbEntry = "Data Source=";
                                icnt++;
                            }

                            if (sKey.ToUpper() == "DATABASE")
                            {
                                sValue = fileData[1];
                                sDbName = sValue;
                                sDbEntry += "Initial Catalog=";
                                icnt++;
                            }

                            if (sKey.ToUpper() == "USERID")
                            {
                                sValue = fileData[1];
                                sDbEntry += "User ID=";
                                icnt++;
                            }

                            if (sKey.ToUpper() == "PASSWORD")
                            {
                                sValue = fileData[1];
                                sDbEntry += "Password=";
                                icnt++;
                            }

                            sDbEntry += sValue;
                            if (icnt != 4)
                                sDbEntry += ";";

                            if (icnt == 4)
                            {
                                m_oDbConns.Add(sDbName, sDbEntry);
                                sDbName = "";
                                sDbEntry = "";
                                icnt = 0;
                                bRet = true;
                            }
                        }
                       // Data Source=10.192.211.228;Initial Catalog=WS_DB_94_1;User ID=SA;Password=Password1
                    }
                }
                catch (Exception ex)
                {
                    Log.WriteLine("Folder: {0} ", ex.Message);
                }
            }
            return bRet;
            
        }

        public void ScanAllFolderMappings(string[] args)
        {
            if (args.Length < 8)
            {
                Console.WriteLine("Syntax: <Command> <WorkServer> <NRTAdmin> <password> <ImpersonatorSMTP> <password> <exchange server name> <ExchangeVersion>");
                Console.WriteLine("Example: GET_ALL_FOLDER_MAPPINGS WorkSite NRTAdmin password ImpersonatorSMTPAddress@dev.local password exchangeServer Exchange2010");
                //GET_ALL_FOLDER_MAPPINGS 10.192.211.228 ewsuser mhdocs ewsuser@exdev2016.local !manage6 10.192.211.238 Exchange2010_SP1
                // SCAN_SEARCH_FOLDER ewsuser@exdev2016.local !manage6 10.192.211.238 Exchange2010_SP1 1 1 Queued False
                return;
            }
            do
            {

                if (!File.Exists("Users.txt"))
                {
                    Console.WriteLine("Users.txt doesn't exist");
                    Log.WriteLine("Users.txt doesn't exist");
                    return;
                }

                string sImpersonationAC = args[4];
                string sPassword = args[5];
                string sExchServer = args[6];
                string sExchVersion = args[7];
                bool bHeaderAdded = false;

                
                Log.AutoFlush = true;
                EwsMappedFolderReportLog.AutoFlush = true;

                System.IO.StreamReader file = new System.IO.StreamReader("Users.txt");
                string line;
                
                m_oQueuedEmails = new Dictionary<String, List<ExchangeQueuedEmails>>();


                while ((line = file.ReadLine()) != null)
                {
                    try
                    {
                        if (line.Length > 0)
                        {

                            Log.WriteLine("=========================================================================================================================================================");
                            Console.WriteLine("");
                            Console.WriteLine("=============================================");
                            Log.WriteLine("Processing {0}", line);
                            Console.WriteLine("Processing {0}", line);
                            ExchangeService service;

                            ExchangeVersion exchVer = ExchangeVersion.Exchange2010;
                            if (sExchVersion == "")
                                exchVer = ExchangeVersion.Exchange2010;
                            if (sExchVersion == "Exchange2010")
                                exchVer = ExchangeVersion.Exchange2010;
                            if (sExchVersion == "Exchange2010_SP1")
                                exchVer = ExchangeVersion.Exchange2010_SP1;
                            if (sExchVersion == "Exchange2010_SP2")
                                exchVer = ExchangeVersion.Exchange2010_SP2;
                            if (sExchVersion == "Exchange2007_SP1")
                                exchVer = ExchangeVersion.Exchange2007_SP1;
                            if (sExchVersion == "Exchange2013")
                                exchVer = ExchangeVersion.Exchange2013;
                            if (sExchVersion == "Exchange2013_SP1")
                                exchVer = ExchangeVersion.Exchange2013_SP1;
                            if (sExchVersion == "Exchange2016")
                                exchVer = ExchangeVersion.Exchange2013;

                            service = new ExchangeService(exchVer);

                            service.Credentials = new WebCredentials(sImpersonationAC, sPassword);
                            service.TraceListener = new TraceListener();
                            service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

                            string[] fileData = line.Split(':');
                            string smtpAddress = "";
                            string sUser = "";
                            if (fileData.Count() > 1)
                            {
                                smtpAddress = fileData[0];
                                sUser = fileData[1];
                            }


                            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);


                            String sExchSrv;
                            string[] exchArr = sExchServer.Split('>');
                            if (exchArr.Count() > 1)
                                sExchSrv = exchArr[1];
                            else if (sExchServer.Length > 0)
                                sExchSrv = sExchServer;
                            else
                            {
                                Log.WriteLine("Exchange server field is blank");
                                Console.WriteLine("Exchange server field is blank");
                                break;
                            }

                            string exchangeUrl;
                            exchangeUrl = "https://";
                            exchangeUrl += sExchSrv;
                            exchangeUrl += "/EWS/Exchange.asmx";


                            service.Url = new Uri(exchangeUrl);
                            
                            ServicePointManager.ServerCertificateValidationCallback = Program.CertificateValidationCallback;

                            service.TraceEnabled = true;

                            m_oWorkSession.ClearEmAndMappingRequests();

                            if (!bHeaderAdded)
                            {
                                EwsMappedFolderReportLog.WriteLine("Database, SID, Work FolderID, USERID, Status, Status Description, Enabled, Operator, Last Sync, DMS Folder, Outlook Folder, EntryID");
                                bHeaderAdded = true;
                            }

                            if (m_oWorkSession.GetMappedFolders(sUser))
                            {
                                Dictionary<String, FolderMapping> oEmFolderMappings = null;

                                m_oWorkSession.GetMappedFolderCollection(ref oEmFolderMappings);

                                if (oEmFolderMappings != null)
                                {
                                    string sMsg = "";
                                    string foldName = "";
                                    foreach (KeyValuePair<String, FolderMapping> mapping in oEmFolderMappings)
                                    {
                                        try
                                        {
                                            foldName = "";
                                            sMsg = "";
                                            String FoldEwsId;
                                            FoldEwsId = ConvertID(ref service, smtpAddress, "HEX", "EWSID", mapping.Value.exchFolderID);
                                            if ((FoldEwsId != null) && (FoldEwsId.Length > 0))
                                            {
                                                Folder fld;
                                                FolderId id = new FolderId(FoldEwsId);

                                                fld = Folder.Bind(service, id);
                                                if (fld != null)
                                                {
                                                    foldName = fld.DisplayName;
                                                    foldName = foldName.Replace(',', '_');
                                                }
                                            }
                                            sMsg += mapping.Value.databaseName;
                                            sMsg += ",";
                                            sMsg += mapping.Value.sid;
                                            sMsg += ",";
                                            sMsg += mapping.Value.prjID.ToString();
                                            sMsg += ",";
                                            sMsg += mapping.Value.userID;
                                            sMsg += ",";
                                            //sMsg += mapping.Value.login;
                                            //sMsg += ",";
                                            sMsg += mapping.Value.status.ToString();
                                            sMsg += ",";

                                            string sDesc = mapping.Value.statusDescription;
                                            if (sDesc != null)
                                                sDesc = sDesc.Replace(',', '_');
                                            else
                                                sDesc = "";

                                            sMsg += sDesc;
                                            sMsg += ",";
                                            sMsg += mapping.Value.isActive.ToString();
                                            sMsg += ",";
                                            sMsg += mapping.Value.sOperator;
                                            sMsg += ",";
                                            sMsg += mapping.Value.lastSyncTime;
                                            sMsg += ",";
                                            string sFoldPath = mapping.Value.projectFolderPath;
                                            if (sFoldPath != null)
                                                sFoldPath = sFoldPath.Replace(',', '_');
                                            else
                                                sFoldPath = "";
                                            sMsg += sFoldPath;
                                            sMsg += ",";

                                            string sFoldName = foldName;
                                            if (sFoldName != null)
                                                sFoldName = sFoldName.Replace(',', '_');
                                            else
                                                sFoldName = "";
                                            sMsg += sFoldName;// foldName;
                                            sMsg += ",";
                                            sMsg += mapping.Value.foldEntryId;
                                            EwsMappedFolderReportLog.WriteLine(sMsg);
                                        }
                                        catch (Exception ex)
                                        {
                                            DateTime dt = DateTime.Now;
                                            //String.Format("{0:u}", dt);
                                            Log.WriteLine("ScanAllFolderMapping report: {0}, Message: {1} EntryId: {2}, SID: {3} ", dt, ex.Message, mapping.Key, mapping.Value.sid);
                                            Log.WriteLine(" ");
                                        }
                                    }
                                }
                            }

                        }
                    }
                    catch (Exception ex)
                    {
                        Log.WriteLine("Folder: {0} ", ex.Message);
                    }

                }

            } while (false);
        }

        public String ConvertID(ref ExchangeService service, string sSmtpAddr, string sSourceIdType, string sDestIdType, string sId)
        {
            AlternateId oAltID = new AlternateId();
            if (sSourceIdType.ToUpper() == "HEX")
                oAltID.Format = IdFormat.HexEntryId;
            else
                oAltID.Format = IdFormat.EwsId;
            oAltID.Mailbox = sSmtpAddr;
            oAltID.UniqueId = sId;

            IdFormat destIdFormat;
            if (sDestIdType.ToUpper() == "HEX")
                destIdFormat = IdFormat.HexEntryId;
            else
                destIdFormat = IdFormat.EwsId;

            //Convert  PR_ENTRYID identifier format to an EWS identifier. 
            AlternateIdBase oAltIDBase = service.ConvertId(oAltID, destIdFormat);
            AlternateId oAltIDResp = (AlternateId)oAltIDBase;

            return oAltIDResp.UniqueId; //Entry.Key;
        }
        //Scan all the email with 'Queued' Status within the search folder. Filing status can be changed in the cmd line arg
        public void ScanSearchFolderForEmails(string[] args)
        {
            if (args.Length < 13)
            {
                Console.WriteLine("Syntax: <Command> <WorkServer> <NRTAdmin> <password> <ImpersonatorSMTP> <password> <exchange server name> <ExchangeVersion> <SearchFolderParent> <SearchFolder> <FilingStatus> <CountOnly> <ReportMode>");
                Console.WriteLine("Example: SCAN_SEARCH_FOLDER WorkSite NRTAdmin password ImpersonatorSMTPAddress@dev.local password exchangeServer Exchange2010 2 1 Queued False True");
                //SCAN_SEARCH_FOLDER 10.192.211.228 ewsuser mhdocs ewsuser@exdev2016.local !manage6 10.192.211.238 Exchange2010_SP1 2 1 Queued False True
                return;
            }

            do
            {

                if (!File.Exists("Users.txt"))
                {
                    Console.WriteLine("Users.txt doesn't exist");
                    Log.WriteLine("Users.txt doesn't exist");
                    return;
                }

                string sImpersonationAC = args[4];
                string sPassword = args[5];
                string sExchServer = args[6];
                string sExchVersion = args[7];
                string sSearchFoldParent = args[8];
                string sSearchFold = args[9];
                string sSearchFor = args[10];
                string sCountOnly = args[11];
                string sReportMode = args[12].ToUpper();

                if (sSearchFor.ToUpper() == "QUEUED")
                    sSearchFor = "Queued";
                if (sSearchFor.ToUpper() == "ERROR")
                    sSearchFor = "Error";
                if (sSearchFor.ToUpper() == "FILED")
                    sSearchFor = "Filed";

                bool bCountOnly = false;
                if (sCountOnly.ToUpper() == "TRUE")
                    bCountOnly = true;

                Log.AutoFlush = true;
                EwsReportLog.AutoFlush = true;

                System.IO.StreamReader file = new System.IO.StreamReader("Users.txt");
                string line;
                long iTotalEmailCount = 0;
                bool bHeaderAdded = false;

                m_oQueuedEmails = new Dictionary<String, List<ExchangeQueuedEmails>>();


                while ((line = file.ReadLine()) != null)
                {
                    try
                    {
                        if (line.Length > 0)
                        {

                            Log.WriteLine("=========================================================================================================================================================");
                            Console.WriteLine("");
                            Console.WriteLine("=============================================");
                            Log.WriteLine("Processing {0}", line);
                            Console.WriteLine("Processing {0}", line);
                            ExchangeService service;
                            
                            ExchangeVersion exchVer = ExchangeVersion.Exchange2010;
                            if (sExchVersion == "")
                                exchVer = ExchangeVersion.Exchange2010;
                            if (sExchVersion == "Exchange2010")
                                exchVer = ExchangeVersion.Exchange2010;
                            if (sExchVersion == "Exchange2010_SP1")
                                exchVer = ExchangeVersion.Exchange2010_SP1;
                            if (sExchVersion == "Exchange2010_SP2")
                                exchVer = ExchangeVersion.Exchange2010_SP2;
                            if (sExchVersion == "Exchange2007_SP1")
                                exchVer = ExchangeVersion.Exchange2007_SP1;
                            if (sExchVersion == "Exchange2013")
                                exchVer = ExchangeVersion.Exchange2013;
                            if (sExchVersion == "Exchange2013_SP1")
                                exchVer = ExchangeVersion.Exchange2013_SP1;
                            if (sExchVersion == "Exchange2016")
                                exchVer = ExchangeVersion.Exchange2013;

                            service = new ExchangeService(exchVer);

                            service.Credentials = new WebCredentials(sImpersonationAC, sPassword);
                            service.TraceListener = new TraceListener();
                            service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

                            string[] fileData = line.Split(':');
                            string smtpAddress = "";
                            string sUser = "";
                            if (fileData.Count() > 1)
                            {
                                smtpAddress = fileData[0];
                                sUser = fileData[1];
                            }

                           
                            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

                           
                            String sExchSrv;
                            string[] exchArr = sExchServer.Split('>');
                            if (exchArr.Count() > 1)
                                sExchSrv = exchArr[1];
                            else if (sExchServer.Length > 0)
                                sExchSrv = sExchServer;
                            else
                            {
                                Log.WriteLine("Exchange server field is blank");
                                Console.WriteLine("Exchange server field is blank");
                                break;
                            }

                            string exchangeUrl;
                            exchangeUrl = "https://";
                            exchangeUrl += sExchSrv;
                            exchangeUrl += "/EWS/Exchange.asmx";


                            service.Url = new Uri(exchangeUrl);


                            ServicePointManager.ServerCertificateValidationCallback = Program.CertificateValidationCallback;

                            service.TraceEnabled = true;

                            m_oWorkSession.ClearEmAndMappingRequests();

                            FolderView folderView = new FolderView(5);
                            folderView.PropertySet = new PropertySet(FolderSchema.DisplayName);


                            SearchFilter searchFoldFilter = null;
                            string sFolderName = "";
                        

                            if (sSearchFold == "1")
                            {
                                sFolderName = "WCSE_FolderMappings";
                                searchFoldFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, sFolderName);
                            }
                            else if (sSearchFold == "2")
                            {
                                sFolderName = "WCSE_SFMailboxSync";
                                searchFoldFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, sFolderName);
                            }

                            WellKnownFolderName wellknownFoldName;
                            if (sSearchFoldParent == "1")
                                wellknownFoldName = WellKnownFolderName.Root;
                            else if (sSearchFoldParent == "2")
                                wellknownFoldName = WellKnownFolderName.MsgFolderRoot;
                            else if (sSearchFoldParent == "3")
                                wellknownFoldName = WellKnownFolderName.SearchFolders;
                            else
                                wellknownFoldName = WellKnownFolderName.Root;

                            FindFoldersResults findFoldResults = service.FindFolders(wellknownFoldName, searchFoldFilter, folderView);

                            Dictionary<String, String> oParentFolders = null;
                            oParentFolders = new Dictionary<String, String>();

                            ExtendedPropertyDefinition PidTagInternetMessageId = new ExtendedPropertyDefinition(4149, MapiPropertyType.String);

                            ExtendedPropertyDefinition PidTagSearchKey = new ExtendedPropertyDefinition(12299, MapiPropertyType.Binary);

                            ExtendedPropertyDefinition filingStatus = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                                     "FilingStatus", MapiPropertyType.String);

                            List<ExchangeQueuedEmails> userQueuedEmails = new List<ExchangeQueuedEmails>();

                            m_oWorkSession.GetExplicitRequests(sUser, EMRequestStatus.EMRequestSubmitted);
                            m_oWorkSession.GetExplicitRequests(sUser, EMRequestStatus.EMRequestFailure);
                            m_oWorkSession.GetExplicitRequests(sUser, EMRequestStatus.EMRequestToBeStamped);
                            m_oWorkSession.GetMappedFolders(sUser);

                            foreach (Folder folder in findFoldResults.Folders)
                            {
                                if (folder is SearchFolder && folder.DisplayName.Equals(sFolderName))
                                {
                                    
                                    FindItemsResults<Item> findResults;
                                    ItemView view = new ItemView(50, 0, OffsetBasePoint.Beginning);                                   
                                    view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
                                    view.Traversal = ItemTraversal.Shallow;
                                    

                                    SearchFilter.SearchFilterCollection searchOrFilterCollection = new SearchFilter.SearchFilterCollection(LogicalOperator.Or);

                                    searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(filingStatus, sSearchFor));
                                    searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.WorkSite.Ems.Queued"));
                                    //SearchFilter searchFlt = new SearchFilter.IsEqualTo(filingStatus, sSearchFor);
                                    do
                                    {
                                        int iTotalUpdates = 0;
                                        Dictionary<String, ExchangeQueuedEmails> m_oEmRequestBucket = new Dictionary<String, ExchangeQueuedEmails>();
                                        Dictionary<String, ExchangeQueuedEmails> m_oNoEmRequestBucket = new Dictionary<String, ExchangeQueuedEmails>();                                        

                                        // Send the request to search the Inbox and get the results.
                                        findResults = service.FindItems(folder.Id, searchOrFilterCollection, view);

                                        if (bCountOnly)
                                        {
                                            iTotalEmailCount = findResults.TotalCount;
                                            break;
                                        }
                                        else
                                        {
                                            foreach (Item myItem in findResults.Items)
                                            {
                                                PropertySet propEmail = new PropertySet(ItemSchema.Subject, ItemSchema.ItemClass,
                                                                                        ItemSchema.LastModifiedTime, ItemSchema.ParentFolderId,
                                                                                        PidTagInternetMessageId, PidTagSearchKey);

                                                Item item = Item.Bind(service, myItem.Id.UniqueId, propEmail);

                                                Folder fld = null;
                                                string sDispName = "";
                                                if (!oParentFolders.ContainsKey(item.ParentFolderId.UniqueId))
                                                {
                                                    fld = Folder.Bind(service, item.ParentFolderId.UniqueId);
                                                    oParentFolders.Add(item.ParentFolderId.UniqueId, fld.DisplayName);
                                                    sDispName = fld.DisplayName;
                                                }
                                                else
                                                    oParentFolders.TryGetValue(item.ParentFolderId.UniqueId, out sDispName);

                                                
                                                iTotalEmailCount++;

                                                ExchangeQueuedEmails queuedEmail = new ExchangeQueuedEmails();
                                                queuedEmail.User = sUser;
                                                queuedEmail.EmailId = smtpAddress;
                                                queuedEmail.subject = item.Subject;
                                                queuedEmail.lastModifiedTime = item.LastModifiedTime.ToString();
                                                queuedEmail.messageClass = item.ItemClass;
                                                queuedEmail.ewsId = item.Id.UniqueId;
                                                

                                                AlternateId oAltID = new AlternateId();
                                                oAltID.Format = IdFormat.EwsId;
                                                oAltID.Mailbox = smtpAddress;
                                                oAltID.UniqueId = item.Id.UniqueId;
    
                                                AlternateIdBase oAltIDBase = service.ConvertId(oAltID, IdFormat.HexEntryId);
                                                AlternateId oAltIDResp = (AlternateId)oAltIDBase;
                                                queuedEmail.entryId = oAltIDResp.UniqueId;

                                                

                                                queuedEmail.parentFolderEWSId = item.ParentFolderId.UniqueId;
                                                AlternateId oAltIDFold = new AlternateId();
                                                oAltIDFold.Format = IdFormat.EwsId;
                                                oAltIDFold.Mailbox = smtpAddress;
                                                oAltIDFold.UniqueId = item.ParentFolderId.UniqueId;

                                                AlternateIdBase oAltIDBase1 = service.ConvertId(oAltIDFold, IdFormat.HexEntryId);
                                                AlternateId oAltIDResp1 = (AlternateId)oAltIDBase1;
                                                queuedEmail.parentFolderEntryId = oAltIDResp1.UniqueId;


                                                queuedEmail.parentFolderName = sDispName;
                                                //queuedEmail.sentDate = item.DateTimeSent.ToString();   

                                                bool bInternalError = false;
                                                foreach (ExtendedProperty extendedProperty in item.ExtendedProperties)
                                                {
                                                    if (extendedProperty.PropertyDefinition.Tag == 4149)
                                                    {
                                                        queuedEmail.messageId = extendedProperty.Value.ToString();
                                                        if (queuedEmail.messageId.Length > 3)
                                                            queuedEmail.messageId = queuedEmail.messageId.Substring(1, queuedEmail.messageId.Length - 2);
                                                        Console.WriteLine(extendedProperty.Value.ToString());
                                                    }

                                                    if (extendedProperty.PropertyDefinition.Tag == 12299)
                                                    {
                                                        Byte[] searchKeyValue;
                                                        string hexSearchKey = "";
                                                        if (extendedProperty.Value != null)
                                                        {
                                                            searchKeyValue = (Byte[])extendedProperty.Value;
                                                            if (searchKeyValue == null)
                                                            {
                                                                Log.WriteLine("Couldn't get search key for {0}", item.Id.UniqueId);
                                                                propEmail = null;
                                                                bInternalError = true;
                                                            }
                                                            hexSearchKey = BitConverter.ToString(searchKeyValue).Replace("-", "");
                                                            queuedEmail.searchKey = hexSearchKey;
                                                            
                                                        }
                                                    }

                                                }

                                                if (bInternalError)
                                                    continue;

                                                ExplicitRequest exReq;
                                                if (m_oWorkSession.IsEmailRequestExists(queuedEmail, out exReq))
                                               {

                                                   queuedEmail.explicitReq = exReq;
                                                   if (!m_oEmRequestBucket.ContainsKey(queuedEmail.entryId))
                                                   {
                                                       m_oEmRequestBucket.Add(queuedEmail.entryId, queuedEmail);
                                                   }
                                               }
                                               else
                                               {
                                                   if (!m_oNoEmRequestBucket.ContainsKey(queuedEmail.entryId))
                                                   {
                                                       m_oNoEmRequestBucket.Add(queuedEmail.entryId, queuedEmail);
                                                   }
                                               }
                                                
                                            }
                                        }

                                        
                                        ProcessEmRequestBucket(ref m_oEmRequestBucket);
                                        ProcessEmRequestBucket(ref m_oNoEmRequestBucket);

                                        AddRequestsForNotFiledNotQueuedInWSEmails(sReportMode, ref m_oNoEmRequestBucket);

                                        if (!bHeaderAdded)
                                        {
                                            EwsReportLog.WriteLine("Queued In Exchange, Queued In Work Server, Filed In Work Server, Active Request, User, Email ID, Email EntryID, Email EWSID, Subject, Message Class, MessageID, Last Modified Time, Folder Name, Mapped Folder, Folder EntryID, Folder EWSID, Status Description");
                                            bHeaderAdded = true;
                                        }

                                        MarkEmailAsFiledElseQueue(ref service, sReportMode, ref iTotalUpdates, ref m_oNoEmRequestBucket, ref Log);

                                        GenerateReport(1, ref m_oEmRequestBucket);
                                        GenerateReport(2, ref m_oNoEmRequestBucket);

                                        m_oEmRequestBucket = null;
                                        m_oNoEmRequestBucket = null;

                                        if (sReportMode == "TRUE")
                                            view.Offset += 50;
                                        else
                                            view.Offset += 50 - iTotalUpdates;
                                    } while (findResults.MoreAvailable);

                                    Log.WriteLine("Total emails in {0} status for {1} - {2}", sSearchFor, smtpAddress, iTotalEmailCount);
                                    Log.WriteLine("------------------------------------------------------------------------------------");
                                    break;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        DateTime dt = DateTime.Now;
                        Log.WriteLine("Folder: {0} - {1} ", dt, ex.Message);
                    }    

                    //GenerateReport(1, ref m_oEmRequestBucket);
                    //GenerateReport(2, ref m_oNoEmRequestBucket);
                }
                
            } while (false);


        }


        public bool MarkEmailAsFiled(ref ExchangeService service, string sReportMode, string ewsId, double DocNum, int Version, ref string sDb, ref int iPrjId, ref int iTotalUpdates, ref StreamWriter Log)
        {
            bool bRet = false;

            try
            {
                if (DocNum <= 0)
                {
                    Log.WriteLine("MarkEmailAsFiled: Invalid parameter DocNum: {0}, Version: {1}", DocNum, Version);
                    return bRet;
                }

                if (Version <= 0)
                    Version = 1;


                ExtendedPropertyDefinition emailCount = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                       "FilingCount", MapiPropertyType.Long);

                ExtendedPropertyDefinition filingStatusCode = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                        "FilingStatusCode", MapiPropertyType.Integer);

                ExtendedPropertyDefinition emailFilingStatus = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                        "FilingStatus", MapiPropertyType.String);

                ExtendedPropertyDefinition filingDocumentId = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                        "FilingDocumentID", MapiPropertyType.String);

                ExtendedPropertyDefinition filingFolder = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                        "FilingFolder", MapiPropertyType.String);

                ExtendedPropertyDefinition filingDate = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                        "FilingDate", MapiPropertyType.SystemTime);

                ExtendedPropertyDefinition filingLocation = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                        "FilingLocation", MapiPropertyType.String);

                ExtendedPropertyDefinition autnLastChangeTime = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                        "x-autn-lastchange-time", MapiPropertyType.SystemTime);


                string sPath = "";
                string sServerName = "";
                sDb = "";
                iPrjId = 0;
                if (!m_oWorkSession.GetWhereFiledInfo(DocNum, Version, ref sPath, ref sServerName, ref sDb, ref iPrjId, ref Log))
                {
                    Log.WriteLine("MarkEmailAsFiled: GetWhereFiledInfo failed");
                    return bRet;
                }

                if ((sPath == String.Empty) || (sServerName == String.Empty) || (sDb == String.Empty))
                {
                    Log.WriteLine("MarkEmailAsFiled: couldn't retrieve enough data. Path: {0}, Server: {1}, Database: {2}", sPath, sServerName, sDb);
                    return bRet;
                }


                var bindResults = service.BindToItems(new[] { new ItemId(ewsId) }, new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.ItemClass));//,ItemSchema.Subject, ItemSchema.ItemClass, emailFilingStatus));
                foreach (GetItemResponse getItemResponse in bindResults)
                {
                    Item item = getItemResponse.Item;
                    if (item != null)
                    {

                        item.SetExtendedProperty(emailCount, "1");
                        item.SetExtendedProperty(filingStatusCode, "2");
                        item.SetExtendedProperty(emailFilingStatus, "Filed");


                        string sMoniker = "";
                        sMoniker = "!nrtdms:0:!session:";
                        sMoniker += sServerName;
                        sMoniker += ":!database:";
                        sMoniker += sDb;
                        sMoniker += ":!document:";
                        sMoniker += DocNum.ToString();
                        sMoniker += ",";
                        sMoniker += Version.ToString();
                        sMoniker += ":";

                        item.SetExtendedProperty(filingDocumentId, sMoniker);
                        item.SetExtendedProperty(filingFolder, sPath);
                        item.SetExtendedProperty(filingDate, DateTime.Now.ToString());
                        item.SetExtendedProperty(filingLocation, sPath);
                        item.SetExtendedProperty(autnLastChangeTime, DateTime.Now.ToString());

                        if (item.ItemClass == "IPM.Note.WorkSite.Ems.Queued")
                            item.ItemClass = "IPM.Note.WorkSite.Ems.Filed";

                        if (sReportMode == "FALSE")
                        {
                            item.Update(ConflictResolutionMode.AlwaysOverwrite);
                            iTotalUpdates++;
                        }

                        Log.WriteLine("Updated email: {0} ", item.Subject);

                        bRet = true;

                    }
                }
            }
            catch (Exception ex)
            {
                Log.WriteLine("MarkEmailAsFiled failed :{0} ", ex.Message);
            }
            return bRet;

        }

       
        public bool MarkEmailAsFiledElseQueue(ref ExchangeService service, string sReportMode, ref int iTotalUpdates, ref Dictionary<String, ExchangeQueuedEmails> oBucket, ref StreamWriter Log)
        {
            bool bRet = false;

            string sDb = "";
            int iPrjId = 0;
            foreach (KeyValuePair<String, ExchangeQueuedEmails> entry in oBucket)
            {               
                if ((entry.Value == null) || (entry.Value.iExistInWorkServer != 1) || (entry.Value.ewsId.Length == 0) || (entry.Value.ewsId.Length == 0))
                    continue;

                sDb = "";
                if (!MarkEmailAsFiled(ref service, sReportMode, entry.Value.ewsId, 
                                        entry.Value.filedEmailDetails.DocNum, 
                                        entry.Value.filedEmailDetails.Version, ref sDb, ref iPrjId, ref iTotalUpdates, ref Log))
                {
                    Log.WriteLine("MarkEmailAsFiled failed EWSID:{0}, EntryID: {1} ", entry.Value.ewsId, entry.Value.entryId);
                    entry.Value.PrjId = iPrjId;
                    ExchangeQueuedEmails queuedEmail = entry.Value;
                    m_oWorkSession.InsertEMRequestEntry(sDb, queuedEmail, ref Log);
                }

                

            }

            return bRet;
        }

        public void GenerateReport(int iBucket, ref Dictionary<String, ExchangeQueuedEmails> oBucket)
        {
            
            string sMsg = "";
          

            // Queued In Exchange, Queued In Work, Filed In Work, Status, Email EntryID, Email EWSID, Subject, Message Class, MessageID, Last Modified Date, Folder EntryID, Folder EWSID, Folder Name,
               foreach (KeyValuePair<String, ExchangeQueuedEmails> entry in oBucket)
                {
                    try
                    {
                       

                        if (iBucket == 1)
                            sMsg = "Y,QUEUED IN WS,";
                        else if (iBucket == 2)
                            sMsg = "Y,NOT QUEUED IN WS,";

                        if (entry.Value.iExistInWorkServer == 1)
                            sMsg += "FILED IN WS";
                        else if (entry.Value.iExistInWorkServer == 0)
                            sMsg += "NOT FILED IN WS";
                        else if (entry.Value.iExistInWorkServer == 2)
                            sMsg += "UNKNOWN";
                        //sMsg += entry.Value.bExistInWorkServer.ToString().ToUpper(); // Filed In Work
                        sMsg += ",";
                        if (entry.Value.explicitReq != null)
                            sMsg += entry.Value.explicitReq.IsActive.ToString().ToUpper(); //Status
                        else
                            sMsg += "FALSE";
                        sMsg += ",";
                        string sUser = entry.Value.User;
                        if (sUser != null)
                            sUser = sUser.Replace(',', '_');
                        else
                            sUser = "";
                        sMsg += sUser;
                        sMsg += ",";
                        sMsg += entry.Value.EmailId;
                        sMsg += ",";
                        sMsg += entry.Value.entryId.ToString(); //EntryId
                        sMsg += ",";
                        sMsg += entry.Value.ewsId.ToString(); // EWSId
                        sMsg += ",";
                        string sSub = entry.Value.subject;
                        if (sSub != null)
                            sSub = sSub.Replace(',', '_');
                        else
                            sSub = "";
                        sMsg += sSub; //Subject
                        sMsg += ",";
                        sMsg += entry.Value.messageClass; //Message class
                        sMsg += ",";
                        sMsg += entry.Value.messageId; // Message ID
                        sMsg += ",";
                        sMsg += entry.Value.lastModifiedTime; // LastModified Time
                        sMsg += ",";

                        string sParentFoldName = entry.Value.parentFolderName;
                        if (sParentFoldName != null)
                            sParentFoldName = sParentFoldName.Replace(',', '_');
                        else
                            sParentFoldName = "";
                        sMsg += sParentFoldName;

                        //sMsg += entry.Value.parentFolderName; // Parent Folder Name
                        sMsg += ",";
                        //if (entry.Value.explicitReq != null)
                        //    sMsg += entry.Value.explicitReq.IsInMappedFolder.ToString().ToUpper(); //Mapped folder
                        //else
                        //    sMsg += "FALSE";
                        FolderMapping foldMapping = null;
                        if (m_oWorkSession.IsFolderMappingExist(entry.Value.parentFolderEntryId, out foldMapping))
                        {
                            if (foldMapping != null)
                            {
                                if (foldMapping != null)
                                    sMsg += "TRUE";
                                else
                                    sMsg += "FALSE";
                            }
                            else
                                sMsg += "FALSE";
                        }
                        else
                            sMsg += "FALSE";

                        sMsg += ",";
                        sMsg += entry.Value.parentFolderEntryId; // Parent EntryID
                        sMsg += ",";
                        sMsg += entry.Value.parentFolderEWSId;
                        sMsg += ",";
                        if (entry.Value.explicitReq != null)
                            sMsg += entry.Value.explicitReq.StatusDescription;
                        else
                            sMsg += "";



                        EwsReportLog.WriteLine(sMsg);
                    }
                    catch (Exception ex)
                    {
                        Log.WriteLine("GenerateReport: {0} ", ex.Message);
                    }
            }
            
            
        }


        public bool AddRequestsForNotFiledNotQueuedInWSEmails(string sReportMode, ref Dictionary<String, ExchangeQueuedEmails> oNoEmRequestBucket)
        {
            bool bRet = false;
            foreach (KeyValuePair<String, ExchangeQueuedEmails> entry in oNoEmRequestBucket)
            {
                if (entry.Value != null)
                {
                    if (entry.Value.iExistInWorkServer == 0)
                    {
                        FolderMapping foldMapping = null;
                        if (m_oWorkSession.IsFolderMappingExist(entry.Value.parentFolderEntryId, out foldMapping))
                        {
                            Log.WriteLine("Making an entry for NotFiledNotQueued item EWSID:{0} ", entry.Value.ewsId);

                            if (foldMapping.prjID <= 0)
                            {
                                Log.WriteLine("Failed to make an entry for the item EWSID:{0} because of invalid prjID: {1} ", entry.Value.ewsId, foldMapping.prjID);
                                continue;
                            }
                            if ((foldMapping.databaseName == null) || (foldMapping.databaseName.Length <= 0))
                            {
                                Log.WriteLine("Failed to make an entry for the item EWSID:{0} because of no database name ", entry.Value.ewsId);
                                continue;
                            }

                            entry.Value.PrjId = foldMapping.prjID;
                            ExchangeQueuedEmails queuedEmail = entry.Value;
                            if (sReportMode == "FALSE")
                                m_oWorkSession.InsertEMRequestEntry(foldMapping.databaseName, queuedEmail, ref Log);
                        }              
                    }
                }
            }

            return bRet;
        }

        public bool ProcessEmRequestBucket(ref Dictionary<String, ExchangeQueuedEmails> oRequestBucket)
        {
            bool bRet = true;


            List<String> msgIdFromExchange = new List<string>();
            List<FiledEmailDetails> msgIdInWorkServer = new List<FiledEmailDetails>();
            List<ExchangeQueuedEmails> EmRequests = new List<ExchangeQueuedEmails>();
            int iCount = 0;
            int iInWSRet = 0;
            bool bFound = false;

            foreach (KeyValuePair<String, ExchangeQueuedEmails> entry in oRequestBucket)
            {
               
                iCount++;
                msgIdFromExchange.Add(entry.Value.messageId);
                EmRequests.Add(entry.Value);
                if (iCount < 5)
                    continue;

                iCount = 0;
                iInWSRet = 0;
                
                
                iInWSRet = m_oWorkSession.CheckEmailExistInWorkServerDatabase(ref m_oDbConns, ref msgIdFromExchange, ref msgIdInWorkServer, ref Log);
                //iInWSRet = m_oWorkSession.CheckEmailExistInWorkServer(ref msgIdFromExchange, ref msgIdInWorkServer, ref Log);

               
                foreach (ExchangeQueuedEmails entry1 in EmRequests)
                {
                    bFound = false;
                    foreach (FiledEmailDetails oValue in msgIdInWorkServer)    
                    {
                        if (entry1.messageId == oValue.messageId)
                        {
                            bFound = true;
                            entry1.filedEmailDetails = oValue;
                        }
                            
                    }

                    if (bFound)
                        entry1.iExistInWorkServer = 1; // Exist
                    else if (iInWSRet == 2)
                        entry1.iExistInWorkServer = 2; // Exception occurred, may be timeout, may be exch conn failed
                    else
                        entry1.iExistInWorkServer = 0; // Doesn't Exist
                }

                

                msgIdFromExchange = null;
                msgIdInWorkServer = null;
                EmRequests = null;

                msgIdFromExchange = new List<string>();
                msgIdInWorkServer = new List<FiledEmailDetails>();
                EmRequests = new List<ExchangeQueuedEmails>();
            }
            
            if (iCount > 0)
            {
                iCount = 0;
                iInWSRet = 0;
                iInWSRet = m_oWorkSession.CheckEmailExistInWorkServerDatabase(ref m_oDbConns, ref msgIdFromExchange, ref msgIdInWorkServer, ref Log); 
                
                foreach (ExchangeQueuedEmails entry1 in EmRequests)
                {
                    bFound = false;
                    foreach (FiledEmailDetails oValue in msgIdInWorkServer)    
                    {
                        if (entry1.messageId == oValue.messageId)
                        {
                            bFound = true;
                            entry1.filedEmailDetails = oValue;
                        }
                    }

                    if (bFound)
                        entry1.iExistInWorkServer = 1;
                    else if (iInWSRet == 2)
                        entry1.iExistInWorkServer = 2;
                    else
                        entry1.iExistInWorkServer = 0;

                }
            }
            msgIdFromExchange = null;
            msgIdInWorkServer = null;
            EmRequests = null;

            return bRet;
        }



        // Get all un-touched (IPM.Note) emails from the mapped folder
        public void GetNotQueuedEmailsFromMappedFolder(string[] args)
        {

            if (args.Length < 5)
            {
                Console.WriteLine("Syntax: <Command> <ImpersonatorSMTP> <password> <end user> <exchange server name> <ExchangeVersion> <MappedFolderFile>");
                Console.WriteLine("Example: GET-MF-NOT-QUEUED-EMAILS ImpersonatorSMTPAddress@dev.local password exchangeServer ExchangeVersion MappedFolder.csv");
               return;
            }
            
            string sImpersonationAC = args[1];
            string sPassword = args[2];
            string sExchServer = args[3];
            string sExchVersion = args[4];
            
            string sMappedFolderFile = args[5];

            Log.AutoFlush = true;
            EwsMFReportLog.AutoFlush = true;

            //("Queued In Exchange, Queued In Work Server, Filed In Work Server, Active Request, User, Email ID, Email EntryID, Email EWSID, Subject, Message Class, MessageID, Last Modified Time, Folder Name, Mapped Folder, Folder EntryID, Folder EWSID, Status Description");

            EwsMFReportLog.WriteLine("User, Subject, Message Class, Email EWSID, Folder Name, Folder EWSID, Total email to reset");
           
            Dictionary<String, String> oFolderEntryIds = null;
            oFolderEntryIds = new Dictionary<String, String>();

            if (!File.Exists(sMappedFolderFile))
            {
                Log.WriteLine("File doesn't exist - {0}", sMappedFolderFile);
                Console.WriteLine("File doesn't exist - {0}", sMappedFolderFile);
                return;
            }

            System.IO.StreamReader file = new System.IO.StreamReader(sMappedFolderFile);

            string line;
            while ((line = file.ReadLine()) != null)
            {
                line.Trim();
                String[] Tokens = line.Split(",".ToCharArray());
                if (2 > Tokens.Length)
                {
                    Log.WriteLine("Invalid entry in {0}", sMappedFolderFile);
                    break;
                }
                if ((Tokens[2] == "N") || (Tokens[3] == "-6"))
                {
                    Log.WriteLine("Disabled: {0}", Tokens[1]);
                    continue;
                }
                if (!oFolderEntryIds.ContainsKey(Tokens[1].ToString()))
                    oFolderEntryIds.Add(Tokens[1].ToString(), Tokens[0].ToString());
                else
                    Console.WriteLine("Record Exist");
            }

            if (oFolderEntryIds.Count() <= 0)
                return;

            ExchangeService service;
            ExchangeVersion exchVer = ExchangeVersion.Exchange2010;
            if (sExchVersion == "")
                exchVer = ExchangeVersion.Exchange2010;
            if (sExchVersion == "Exchange2010")
                exchVer = ExchangeVersion.Exchange2010;
            if (sExchVersion == "Exchange2010_SP1")
                exchVer = ExchangeVersion.Exchange2010_SP1;
            if (sExchVersion == "Exchange2010_SP2")
                exchVer = ExchangeVersion.Exchange2010_SP2;
            if (sExchVersion == "Exchange2007_SP1")
                exchVer = ExchangeVersion.Exchange2007_SP1;
            if (sExchVersion == "Exchange2013")
                exchVer = ExchangeVersion.Exchange2013;
            if (sExchVersion == "Exchange2013_SP1")
                exchVer = ExchangeVersion.Exchange2013_SP1;
            if (sExchVersion == "Exchange2016")
                exchVer = ExchangeVersion.Exchange2013;
            service = new ExchangeService(exchVer);

            service.Credentials = new WebCredentials(sImpersonationAC, sPassword);
            service.TraceListener = new TraceListener();
            service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

            String sExchSrv;
            string[] exchArr = sExchServer.Split('>');
            if (exchArr.Count() > 1)
                sExchSrv = exchArr[1];
            else if (sExchServer.Length > 0)
                sExchSrv = sExchServer;
            else
            {
                Log.WriteLine("Exchange server field is blank");
                Console.WriteLine("Exchange server field is blank");
                return;
            }

            string exchangeUrl;
            exchangeUrl = "https://";
            exchangeUrl += sExchSrv;
            exchangeUrl += "/EWS/Exchange.asmx";

            service.TraceEnabled = true;
            ServicePointManager.ServerCertificateValidationCallback = Program.CertificateValidationCallback;

            service.Url = new Uri(exchangeUrl);

            string sMsg;
            foreach (KeyValuePair<String, String> Entry in oFolderEntryIds)
            {

                service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, Entry.Value);

            
                try
                {
                    // Use ConvertID
                    AlternateId oAltID = new AlternateId();
                    oAltID.Format = IdFormat.HexEntryId;
                    oAltID.Mailbox = Entry.Value;//smtpAddress;
                    oAltID.UniqueId = Entry.Key;

                    //Convert  PR_ENTRYID identifier format to an EWS identifier. 
                    AlternateIdBase oAltIDBase = service.ConvertId(oAltID, IdFormat.EwsId);
                    AlternateId oAltIDResp = (AlternateId)oAltIDBase;

                    String FoldEwsId = oAltIDResp.UniqueId; //Entry.Key;
                    String FoldName = Entry.Value;
            

                    Folder fld;
                    FolderId id = new FolderId(FoldEwsId);

                    fld = Folder.Bind(service, id);
                    Console.WriteLine("Folder Name: " + fld.DisplayName);
                    FoldName = fld.DisplayName;

                        
                    ExtendedPropertyDefinition filingStatus = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                                "FilingStatus", MapiPropertyType.String);



                    FindItemsResults<Item> findResults;
                    ItemView view = new ItemView(100, 0, OffsetBasePoint.Beginning);

                    // Identify the Subject properties to return.
                    // Indicate that the base property will be the item identifier
                    view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.LastModifiedTime,
                                                        ItemSchema.DateTimeSent, ItemSchema.ItemClass, filingStatus);

                    // Set the traversal to shallow. (Shallow is the default option; other options are Associated and SoftDeleted.)
                    view.Traversal = ItemTraversal.Shallow;

                    SearchFilter searchFlt = new SearchFilter.Not(new SearchFilter.Exists(filingStatus));

                    long iTotalEmailCount = 0;
                    do
                    {
                        // Send the request to search the Inbox and get the results.
                        findResults = service.FindItems(id, searchFlt, view);

                        foreach (Item myItem in findResults.Items)
                        {
                            iTotalEmailCount++;
                            sMsg = "";

                            sMsg += Entry.Value; // Folder EntryID
                            sMsg += ",";

                            string sSub = myItem.Subject;
                            sSub = sSub.Replace(',', '_');
                            sMsg += sSub;               //Subject
                            sMsg += ",";
                            
                            sMsg += myItem.ItemClass; // Message Class
                            sMsg += ",";
                            
                            sMsg += myItem.Id.UniqueId; // EWSID
                            sMsg += ",";

                            string sFoldName = FoldName;
                            sFoldName = sFoldName.Replace(',', '_');

                            sMsg += sFoldName; // Folder Name
                            sMsg += ",";

                            sMsg += FoldEwsId; // EWSID
                           // sMsg += ",";


                            EwsMFReportLog.WriteLine(sMsg);

                            //Log.WriteLine("Subject: {0}", myItem.Subject);
                            //Log.WriteLine("ItemClass: {0}", myItem.ItemClass);
                            //Log.WriteLine("EWSID: {0}", myItem.Id.UniqueId);
                            //Log.WriteLine("");
                        }

                        view.Offset += 100;
                    } while (findResults.MoreAvailable);


                    //Console.WriteLine("Reset count : {0} ", iTotalEmailCount);
                    //Console.WriteLine("Folder : {0} : Items Processed : {1}", FoldName, iTotalEmailCount);

                    Log.WriteLine("");
                    Log.WriteLine("Total emails reset for {0} on Folder: {1} : {2} : {3}", Entry.Value, FoldName, FoldEwsId, iTotalEmailCount);
                    Log.WriteLine("----------------------------------------------------------");
            
                }
                catch (Exception ex)
                {
                    Log.WriteLine("Folder: {0} : ", ex.Message);
                }
            }

        }


        public void CreateUnfiledSearchFolder(string[] args)
        {
            if (args.Length < 5)
            {
                Console.WriteLine("Syntax: <Command> <ImpersonatorSMTP> <password> <exchange server name> <ExchangeVersion>");
                Console.WriteLine("Example: CREATE_UNFILED_SEARCH_FOLDER ImpersonatorSMTPAddress@dev.local password exchangeServer Exchange2010");
                //CREATE_UNFILED_SEARCH_FOLDER ewsuser@exdev2016.local !manage6 10.192.211.238 Exchange2010_SP1
                return;
            }

            do
            {

                if (!File.Exists("UsersForSearchFolder.txt"))
                {
                    Console.WriteLine("UsersForSearchFolder.txt doesn't exist");
                    Log.WriteLine("UsersForSearchFolder.txt doesn't exist");
                    return;
                }

                string sImpersonationAC = args[1];
                string sPassword = args[2];
                string sExchServer = args[3];
                string sExchVersion = args[4];               
                //string sReportMode = args[5].ToUpper();

                Log.AutoFlush = true;

                System.IO.StreamReader file = new System.IO.StreamReader("UsersForSearchFolder.txt");
                string line;

                while ((line = file.ReadLine()) != null)
                {
                    try
                    {
                        if (line.Length > 0)
                        {

                            Log.WriteLine("=========================================================================================================================================================");
                            Console.WriteLine("");
                            Console.WriteLine("=============================================");
                            Log.WriteLine("Processing {0}", line);
                            Console.WriteLine("Processing {0}", line);
                            ExchangeService service;

                            ExchangeVersion exchVer = ExchangeVersion.Exchange2010;
                            if (sExchVersion == "")
                                exchVer = ExchangeVersion.Exchange2010;
                            if (sExchVersion == "Exchange2010")
                                exchVer = ExchangeVersion.Exchange2010;
                            if (sExchVersion == "Exchange2010_SP1")
                                exchVer = ExchangeVersion.Exchange2010_SP1;
                            if (sExchVersion == "Exchange2010_SP2")
                                exchVer = ExchangeVersion.Exchange2010_SP2;
                            if (sExchVersion == "Exchange2007_SP1")
                                exchVer = ExchangeVersion.Exchange2007_SP1;
                            if (sExchVersion == "Exchange2013")
                                exchVer = ExchangeVersion.Exchange2013;
                            if (sExchVersion == "Exchange2013_SP1")
                                exchVer = ExchangeVersion.Exchange2013_SP1;
                            if (sExchVersion == "Exchange2016")
                                exchVer = ExchangeVersion.Exchange2013;

                            service = new ExchangeService(exchVer);

                            service.Credentials = new WebCredentials(sImpersonationAC, sPassword);
                            service.TraceListener = new TraceListener();
                            service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

                            string[] fileData = line.Split(':');
                            string smtpAddress = "";
                            string sUser = "";
                            if (fileData.Count() > 1)
                            {
                                smtpAddress = fileData[0];
                                sUser = fileData[1];
                            }

                            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);


                            String sExchSrv;
                            string[] exchArr = sExchServer.Split('>');
                            if (exchArr.Count() > 1)
                                sExchSrv = exchArr[1];
                            else if (sExchServer.Length > 0)
                                sExchSrv = sExchServer;
                            else
                            {
                                Log.WriteLine("Exchange server field is blank");
                                Console.WriteLine("Exchange server field is blank");
                                break;
                            }

                            string exchangeUrl;
                            exchangeUrl = "https://";
                            exchangeUrl += sExchSrv;
                            exchangeUrl += "/EWS/Exchange.asmx";


                            service.Url = new Uri(exchangeUrl);


                            ServicePointManager.ServerCertificateValidationCallback = Program.CertificateValidationCallback;

                            service.TraceEnabled = true;

                            SearchFolder searchFolder1 = new SearchFolder(service);

                            SearchFilter.SearchFilterCollection searchOrFilterCollection = new SearchFilter.SearchFilterCollection(LogicalOperator.Or);
                            ExtendedPropertyDefinition filingStatus = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                                     "FilingStatus", MapiPropertyType.String);

                            searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(filingStatus, "Queued"));
                            searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.WorkSite.Ems.Queued"));

                           
                            searchFolder1.SearchParameters.RootFolderIds.Add(WellKnownFolderName.MsgFolderRoot);
                            searchFolder1.SearchParameters.Traversal = SearchFolderTraversal.Deep;
                            searchFolder1.SearchParameters.SearchFilter = searchOrFilterCollection;
                            searchFolder1.DisplayName = "_iM Unfiled Emails";

                            searchFolder1.Save(WellKnownFolderName.SearchFolders);
                            Log.WriteLine("Created search folder: {0} for : {1}", searchFolder1.DisplayName, smtpAddress);
                            Console.WriteLine(searchFolder1.DisplayName + " added.");
                        }
                    }
                    catch (Exception ex)
                    {
                        DateTime dt = DateTime.Now;
                        Log.WriteLine("Folder: {0} - {1} ", dt, ex.Message);
                    }
                }

            } while (false);           
        }


        public void UpdateMsgClsBasedOnFilingStatus(string[] args)
        {
            
            if (args.Length < 12)
            {
                Console.WriteLine("Syntax: <Command> <ImpersonatorSMTP> <password> <exchange server name> <ExchangeVersion> <SearchFolderParent> <SearchFolder> <FilingStatus> <FilingStatusCode> <MessageClass> <CountOnly> <ReportMode>");
                Console.WriteLine("Example: UPDATE_MSG_CLS_BASED_ON_FILING_STATUS ImpersonatorSMTPAddress@dev.local password exchangeServer Exchange2010 2 1 Filed -1 IPM.Note.WorkSite.Ems.Filed False True");
                //SCAN_SEARCH_FOLDER 10.192.211.228 ewsuser mhdocs ewsuser@exdev2016.local !manage6 10.192.211.238 Exchange2010_SP1 2 1 Queued False True
                return;
            }

            do
            {

                if (!File.Exists("Users.txt"))
                {
                    Console.WriteLine("Users.txt doesn't exist");
                    Log.WriteLine("Users.txt doesn't exist");
                    return;
                }

                string sImpersonationAC = args[1];
                string sPassword = args[2];
                string sExchServer = args[3];
                string sExchVersion = args[4];
                string sSearchFoldParent = args[5];
                string sSearchFold = args[6];
                string sSearchFor = args[7];
                string sFilingStatusCode = args[8];
                string sMessageClass = args[9];
                string sCountOnly = args[10];
                string sReportMode = args[11].ToUpper();

                //if (sSearchFor.ToUpper() == "QUEUED")
                //    sSearchFor = "Queued";
                //if (sSearchFor.ToUpper() == "ERROR")
                //    sSearchFor = "Error";
                if (sSearchFor.ToUpper() == "FILED")
                    sSearchFor = "Filed";
                else
                {
                    if ((sFilingStatusCode.Length == 0) || (sFilingStatusCode != "2"))
                    {
                        Console.WriteLine("Search for {0} is not supported", sSearchFor);
                        Log.WriteLine("Search for {0} is not supported", sSearchFor);
                        break;
                    }
                }

                bool bCountOnly = false;
                if (sCountOnly.ToUpper() == "TRUE")
                    bCountOnly = true;

                Log.AutoFlush = true;
               

                System.IO.StreamReader file = new System.IO.StreamReader("Users.txt");
                string line;
                long iTotalEmailCount = 0;
                

                while ((line = file.ReadLine()) != null)
                {
                    try
                    {
                        if (line.Length > 0)
                        {

                            Log.WriteLine("=========================================================================================================================================================");
                            Console.WriteLine("");
                            Console.WriteLine("=============================================");
                            Log.WriteLine("Processing {0}", line);
                            Console.WriteLine("Processing {0}", line);
                            ExchangeService service;

                            ExchangeVersion exchVer = ExchangeVersion.Exchange2010;
                            if (sExchVersion == "")
                                exchVer = ExchangeVersion.Exchange2010;
                            if (sExchVersion == "Exchange2010")
                                exchVer = ExchangeVersion.Exchange2010;
                            if (sExchVersion == "Exchange2010_SP1")
                                exchVer = ExchangeVersion.Exchange2010_SP1;
                            if (sExchVersion == "Exchange2010_SP2")
                                exchVer = ExchangeVersion.Exchange2010_SP2;
                            if (sExchVersion == "Exchange2007_SP1")
                                exchVer = ExchangeVersion.Exchange2007_SP1;
                            if (sExchVersion == "Exchange2013")
                                exchVer = ExchangeVersion.Exchange2013;
                            if (sExchVersion == "Exchange2013_SP1")
                                exchVer = ExchangeVersion.Exchange2013_SP1;
                            if (sExchVersion == "Exchange2016")
                                exchVer = ExchangeVersion.Exchange2013;

                            service = new ExchangeService(exchVer);

                            service.PreAuthenticate = false;

                            service.Credentials = new WebCredentials(sImpersonationAC, sPassword);
                            service.TraceListener = new TraceListener();
                            service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

                            string[] fileData = line.Split(':');
                            string smtpAddress = "";
                            string sUser = "";
                            if (fileData.Count() > 1)
                            {
                                smtpAddress = fileData[0];
                                sUser = fileData[1];
                            }

                           
                            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

                           
                            String sExchSrv;
                            string[] exchArr = sExchServer.Split('>');
                            if (exchArr.Count() > 1)
                                sExchSrv = exchArr[1];
                            else if (sExchServer.Length > 0)
                                sExchSrv = sExchServer;
                            else
                            {
                                Log.WriteLine("Exchange server field is blank");
                                Console.WriteLine("Exchange server field is blank");
                                break;
                            }

                            string exchangeUrl;
                            exchangeUrl = "https://";
                            exchangeUrl += sExchSrv;
                            exchangeUrl += "/EWS/Exchange.asmx";


                            service.Url = new Uri(exchangeUrl);


                            ServicePointManager.ServerCertificateValidationCallback = Program.CertificateValidationCallback;

                            service.TraceEnabled = true;

                           
                            FolderView folderView = new FolderView(5);
                            folderView.PropertySet = new PropertySet(FolderSchema.DisplayName);


                            SearchFilter searchFoldFilter = null;
                            string sFolderName = "";
                        

                            if (sSearchFold == "1")
                            {
                                sFolderName = "WCSE_FolderMappings";
                                searchFoldFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, sFolderName);
                            }
                            else if (sSearchFold == "2")
                            {
                                sFolderName = "WCSE_SFMailboxSync";
                                searchFoldFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, sFolderName);
                            }

                            WellKnownFolderName wellknownFoldName;
                            if (sSearchFoldParent == "1")
                                wellknownFoldName = WellKnownFolderName.Root;
                            else if (sSearchFoldParent == "2")
                                wellknownFoldName = WellKnownFolderName.MsgFolderRoot;
                            else if (sSearchFoldParent == "3")
                                wellknownFoldName = WellKnownFolderName.SearchFolders;
                            else
                                wellknownFoldName = WellKnownFolderName.Root;

                            FindFoldersResults findFoldResults = service.FindFolders(wellknownFoldName, searchFoldFilter, folderView);

                            ExtendedPropertyDefinition filingStatusCode = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                        "FilingStatusCode", MapiPropertyType.Long);

                            ExtendedPropertyDefinition filingStatus = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                                     "FilingStatus", MapiPropertyType.String);

                            
                           
                            foreach (Folder folder in findFoldResults.Folders)
                            {
                                if (folder is SearchFolder && folder.DisplayName.Equals(sFolderName))
                                {
                                    
                                    FindItemsResults<Item> findResults;
                                    ItemView view = new ItemView(50, 0, OffsetBasePoint.Beginning);
                                    view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.ItemClass);
                                    view.Traversal = ItemTraversal.Shallow;


                                    SearchFilter.SearchFilterCollection searchAndFilterCollection = new SearchFilter.SearchFilterCollection(LogicalOperator.And);
                                    SearchFilter.SearchFilterCollection searchOrFilterCollection = new SearchFilter.SearchFilterCollection(LogicalOperator.Or);
                                    
                                    
                                    searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note"));
                                    searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, ""));
                                    //searchAndFilterCollection.Add(new SearchFilter.IsEqualTo(filingStatus, sSearchFor));
                                    
                                    SearchFilter searchFlt;// = new SearchFilter.IsEqualTo(filingStatus, sSearchFor);

                                    if (sSearchFor.Length > 0)
                                        searchAndFilterCollection.Add(new SearchFilter.IsEqualTo(filingStatus, sSearchFor));
                                        //searchFlt = new SearchFilter.IsEqualTo(filingStatus, sSearchFor);
                                    else if (sFilingStatusCode.Length > 0)
                                        searchAndFilterCollection.Add(new SearchFilter.IsEqualTo(filingStatusCode, sFilingStatusCode));
                                        //searchFlt = new SearchFilter.IsEqualTo(filingStatusCode, sFilingStatusCode);
                                    else
                                        break;

                                    searchAndFilterCollection.Add(searchOrFilterCollection);

                                    do
                                    {
                                        int iTotalUpdates = 0;
                                       
                                        // Send the request to search the Inbox and get the results.
                                        findResults = service.FindItems(folder.Id, searchAndFilterCollection, view);

                                        if (bCountOnly)
                                        {
                                            iTotalEmailCount = findResults.TotalCount;
                                            break;
                                        }
                                        else
                                        {
                                            foreach (Item myItem in findResults.Items)
                                            {

                                                if ((myItem.ItemClass == "IPM.Note") ||
                                                    (myItem.ItemClass == "") ||
                                                    (myItem.ItemClass == null))
                                                {
                                                    myItem.ItemClass = "IPM.Note.WorkSite.Ems.Filed";

                                                    iTotalEmailCount++;

                                                    if (sReportMode == "FALSE")
                                                    {
                                                        myItem.Update(ConflictResolutionMode.AlwaysOverwrite);
                                                        iTotalUpdates++;
                                                    }

                                                    Log.WriteLine("Updated email: {0} ", myItem.Subject);
                                                }

                                                
                                            }
                                        }

                                        if (sReportMode == "TRUE")
                                            view.Offset += 50;
                                        else
                                            view.Offset += 50 - iTotalUpdates;
                                    } while (findResults.MoreAvailable);

                                    Log.WriteLine("Total emails in {0} status for {1} - {2}", sSearchFor, smtpAddress, iTotalEmailCount);
                                    Log.WriteLine("------------------------------------------------------------------------------------");
                                    break;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        DateTime dt = DateTime.Now;
                        Log.WriteLine("Folder: {0} - {1} ", dt, ex.Message);
                    }    

                }
                
            } while (false);


        }

  
    }
}
