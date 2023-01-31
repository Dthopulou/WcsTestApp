using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.Exchange.WebServices.Data;
using System.Net;

namespace EWSTestApp
{
    class ResetEmails
    {
        private ExchangeVersion ExchangeServerVer = ExchangeVersion.Exchange2010;
        //private Dictionary<String, List<String>> m_oMarkedFolders = null;
        private Dictionary<String, Dictionary<String, String>> m_oMarkedFolders = new Dictionary<String, Dictionary<String, String>>();
        private Dictionary<String, String> m_oEntryId = null;
        StreamWriter Log = new StreamWriter("EWSTestAppLog.txt", true);

        public void ScanEmailWithEntryId(string[] args)
        {
            long lItemNotExist = 0;
            try
            {
                do
                {
                    Log.AutoFlush = true;

                    if (args.Length < 7)
                    {
                        Console.WriteLine("");
                        Console.WriteLine("FAILED!!! - Invalid parameters");
                        Console.WriteLine("");
                        Console.WriteLine("Syntax: <Command> <ImpersonatorSMTP> <password> <endUserSMTP> <exchange server name> <User EM_REQUEST CSV filePath> <RunReportMode>");
                        //SCAN-FOLDERS admin2@imanage.microsoftonline.com !wov2014 jsmith@imanage.microsoftonline.com ch1prd0410.outlook.com d:\Resubmit1.csv 2015-03-16
                        Console.WriteLine("Example: SCAN-EMAIL ImpersonatorSMTPAddress@dev.local password endUserSMTPAddress xchange.dev.local c:\\User.csv True");

                        break;
                    }

                    if (!File.Exists(args[5]))
                    {
                        Log.WriteLine("File doesn't exist - {0}", args[5]);
                        break;
                    }

                    System.IO.StreamReader file = new System.IO.StreamReader(args[5]);
                    string line;
                    //bool bProcess = false;
                    Int32 sLineNum = 0;
                    String sEntryId;
                    String sFolderPath;
                    String sUserSMTP;
                    String sServer;
                    String sEmailGuid;
                    
                    ExchangeService service = null;

                    bool bInvalidUserDisplayed = false;

                    while ((line = file.ReadLine()) != null)
                    {
                        line.Trim();
                        if (String.IsNullOrEmpty(line))
                        {
                            //Log.WriteLine("line is empty - {0} - {1}", sInputUser, CSVFilePath);
                            continue;
                        }

                        String[] Tokens = line.Split(",".ToCharArray());
                        if (5 > Tokens.Length)
                        {
                            throw new Exception(String.Format("Invalid entry in {0} at line {1}", args[5], sLineNum));
                        }

                        sEntryId = Tokens[0].ToUpper(); // EM_REQUEST - MSG_ID
                        sEmailGuid = Tokens[1].ToUpper(); // EM_REQUEST - EMAIL_GUID
                        sFolderPath = Tokens[2]; // EM_REQUEST - FOLDER_PATH
                        sUserSMTP = Tokens[3].ToUpper(); // DOCUSER - EMAIL
                        sServer = Tokens[4].ToUpper(); // DOCUSER - EXCH_AUTO_DISC

                        if (!BindOneFolder(ref service, args[1], args[2], sUserSMTP, sEntryId, sServer, args[6], ref lItemNotExist))
                            Log.WriteLine(String.Format("Failed to process Folder - EWSId - {0} : ", sEntryId));
                        //else
                        //    Log.WriteLine(String.Format("Processed Folder - EWSId - {0} : ", sEntryId));
                    }

                    

                }
                while (false);
                //Console.WriteLine("Total emails Not found in the store - {0}", lItemNotExist);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
            }
            finally
            {

            }
        }

        public void Execute(string[] args)
        {
            try
            {
                do
                {
                    Log.AutoFlush = true;

                    if (args.Length < 8)
                    {
                        Console.WriteLine("");
                        Console.WriteLine("FAILED!!! - Invalid parameters");
                        Console.WriteLine("");
                        Console.WriteLine("Syntax: <Command> <ImpersonatorSMTP> <password> <endUserSMTP> <exchange server name> <User EM_REQUEST CSV filePath> <Start date> <RunReportMode>");
                        //SCAN-FOLDERS admin2@imanage.microsoftonline.com !wov2014 jsmith@imanage.microsoftonline.com ch1prd0410.outlook.com d:\Resubmit1.csv 2015-03-16
                        Console.WriteLine("Example: SCAN-FOLDERS ImpersonatorSMTPAddress@dev.local password endUserSMTPAddress xchange.dev.local c:\\User.csv startTime True");
                        
                        break;
                    }

                    //m_oMarkedFolders = new Dictionary<String, List<String>>();
                    m_oEntryId = new Dictionary<String, String>();

                    if (!LoadEMProjects(args[3], args[5], 1)) //1 - extract rows which has MSGID=NULL, EMAIL_GUID=NULL
                    {
                        Console.WriteLine(String.Format("Failed to load file {0}", args[5]));
                        Log.WriteLine(String.Format("Failed to load file {0}", args[5]));
                        break;
                    }

                    Log.WriteLine("SCAN-FOLDERS - Processing Users with NULL MSG_ID and NULL EMAIL_GUID");
                    foreach (KeyValuePair<String, Dictionary<String, String>> Entry in m_oMarkedFolders)
                    {
                        String UserSmtp = Entry.Key;
                        Dictionary<String, String> UserFolders = Entry.Value;

                        Log.WriteLine("\n");
                        Log.WriteLine(String.Format("Processing user {0}\n", UserSmtp));
                        Log.WriteLine("\n***************************************************************************************************\n");
                        foreach (KeyValuePair<String, String> Folder in UserFolders)
                        {
                            Log.WriteLine(String.Format("Folder EWSId - {0} : ", Folder.Key));
                            Log.WriteLine(String.Format("Exchange server - {0} ", Folder.Value));

                            if (!ProcessOneFolder(args[1], args[2], UserSmtp, Folder.Key, Folder.Value, args[6], args[7]))
                                Log.WriteLine(String.Format("Failed to process Folder - EWSId - {0} : ", Folder.Key));
                            else
                                Log.WriteLine(String.Format("Processed Folder - EWSId - {0} : ", Folder.Key));

                            Log.WriteLine("\n***************************************************************************************************\n");

                        }
                        
                        
                    }

                }
                while (false);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
            }
            finally
            {
                
            }
        }

        private bool BindOneFolder(ref ExchangeService service,
                                        String sImpersonatorSMTP,
                                        String sImpersonatorPwd,
                                        String sUserSmtp,
                                        String sEmailEntryId,
                                        String sExchServer,                                        
                                        String sReportMode,
                                        ref long lItemNotExist)
        {
            bool bRet = false;
            String sUserSmtpAddr;
            String sFoldEwsId;
            String sExchSrv;
            String sOutlookFolderName = "";
            long iTotalEmailsReset = 0;
            long iSkippedQueuedEMails = 0;
            bool bEmailExist = false;
            try
            {
                do
                {
                    if ((sUserSmtp.Length == 0) || (sExchServer.Length == 0))
                    {
                        Log.WriteLine("Invalid param {0} - {1}", sUserSmtp, sExchServer);
                        break;
                    }

                   
                  

                    // Get Exchange server name
                    string[] exchArr = sExchServer.Split('>');
                    if (exchArr.Count() > 1)
                        sExchSrv = exchArr[1];
                    else if (sExchServer.Length > 0)
                        sExchSrv = sExchServer;
                    else
                    {
                        Log.WriteLine("Exchange server field is blank");
                        break;
                    }



                    bool bConnected = false;
                    if (service == null)
                        bConnected = ConnectToExchangeServer(sExchSrv, sImpersonatorSMTP, sImpersonatorPwd, sUserSmtp, ref service);
                    else
                        bConnected = true;

                    Log.WriteLine("\n***************************************************************************************************\n");

                    if (!bConnected || (null == service) || (null == service.Url))
                    {
                        Log.WriteLine("Failed to connect to exchange server for user " + sUserSmtp);
                        Console.WriteLine("Failed to connect to exchange server for user " + sUserSmtp);
                        break;
                    }

                    String AutoDiscoverURL = service.Url.ToString();
                    AutoDiscoverURL = AutoDiscoverURL.Trim();

                    if (String.Empty == AutoDiscoverURL)
                    {
                        Console.WriteLine("Failed to get exchange server for user " + sUserSmtp);
                        Log.WriteLine("Failed to get exchange server for user " + sUserSmtp);
                        break;
                    }

                    
                    String sEWSID = String.Empty;
                    String sSub = "";
                    String sProp = "";
                    // Create a request to convert identifiers. 
                    AlternateId objAltID = new AlternateId();
                    objAltID.Format = IdFormat.HexEntryId;
                    objAltID.Mailbox = sUserSmtp;
                    objAltID.UniqueId = sEmailEntryId;

                    //Convert  PR_ENTRYID identifier format to an EWS identifier. 
                    AlternateIdBase objAltIDBase = service.ConvertId(objAltID, IdFormat.EwsId);
                    if (null != objAltIDBase)
                    {
                        AlternateId objAltIDResp = (AlternateId)objAltIDBase;
                        sEWSID = objAltIDResp.UniqueId;
                    }

                    if (sEWSID.Length > 1)
                    {
                        ExtendedPropertyDefinition emailFilingStatus = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                    "FilingStatus", MapiPropertyType.String);

                        var bindResults = service.BindToItems(new[] { new ItemId(sEWSID) }, new PropertySet(BasePropertySet.IdOnly,
                                                                        ItemSchema.Subject, ItemSchema.ItemClass, emailFilingStatus));
                        foreach (GetItemResponse getItemResponse in bindResults)
                        {
                            sSub = "";
                            sProp = "";
                            Item item = getItemResponse.Item;
                            if (item != null)
                            {
                                
                                sSub = item.Subject;
                                Console.WriteLine(sSub);
                                Console.WriteLine(item.ItemClass);
                                foreach (ExtendedProperty extendedProperty in item.ExtendedProperties)
                                {
                                    sProp = extendedProperty.PropertyDefinition.Name.ToString() +" : "+ extendedProperty.Value.ToString();
                                    //Console.WriteLine(extendedProperty.PropertyDefinition.Name.ToString());
                                    Console.WriteLine(sProp);
                                    
                                }

                                Log.WriteLine(" Email Exist Subject - {0} ", sSub);
                                Log.WriteLine(" MessageClass - {0} ", item.ItemClass);
                                Log.WriteLine(" Filing Status - {0} ", sProp);
                                bEmailExist = true;
                            }
                            else
                                lItemNotExist++;
                            //sMimeCont = item.MimeContent.ToString();

                        }
                    }

                    bRet = true;
                } while (false);
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("The specified object was not found in the store"))
                    Log.WriteLine(ex.Message);
                else
                    Log.WriteLine(ex.Message + ex.StackTrace);

                bRet = false;
            }
           
            Log.WriteLine("");
            if (!bEmailExist)
            {
                Log.WriteLine(" {0} - {1}", sUserSmtp, sEmailEntryId);
                Log.WriteLine("Total number of emails not exist in the mailbox - {0}", lItemNotExist);
            }
            
            //Log.WriteLine("Total emails skipped for {0} on Folder: {1} - {2} - {3}", sUserSmtp, sOutlookFolderName, sFolderEwsId, iSkippedQueuedEMails);
           

            return bRet;
        }

        private bool ProcessOneFolder(String sImpersonatorSMTP, 
                                        String sImpersonatorPwd, 
                                        String sUserSmtp, 
                                        String sFolderEwsId, 
                                        String sExchServer,
                                        String sStartDate,
                                        String sReportMode )
        {
            bool bRet = false;
            String sUserSmtpAddr;
            String sFoldEwsId;
            String sExchSrv;
            String sOutlookFolderName="";
            long iTotalEmailsReset = 0;
            long iSkippedQueuedEMails = 0;
            try
            {
                do
                {
                    if ((sUserSmtp.Length == 0) || (sFolderEwsId.Length == 0) || (sExchServer.Length == 0))
                    {
                        Log.WriteLine("Invalid param {0} - {1} - {2}", sUserSmtp, sFolderEwsId, sExchServer);
                        break;
                    }

                    if (!sFolderEwsId.Contains("EwsID:")) // Check this
                    {
                        Log.WriteLine("Improper folder id {0}", sFolderEwsId);
                        break;
                    }

                    // Get Folder EWS ID
                    sUserSmtpAddr = sUserSmtp;
                    string[] foldArr = sFolderEwsId.Split(':');
                    if (foldArr.Count() <= 1)
                    {
                        Log.WriteLine("Improper folder-id {0}", sFolderEwsId);
                        break;
                    }
                    sFoldEwsId = foldArr[1];

                    // Get Exchange server name
                    string[] exchArr = sExchServer.Split('>');
                    if (exchArr.Count() > 1)
                        sExchSrv = exchArr[1];
                    else if (sExchServer.Length > 0)
                        sExchSrv = sExchServer;
                    else
                    {
                        Log.WriteLine("Exchange server field is blank");
                        break;
                    }


                    ExchangeService service = null;
                    bool bConnected = ConnectToExchangeServer(sExchSrv, sImpersonatorSMTP, sImpersonatorPwd, sUserSmtp, ref service);

                    Log.WriteLine("\n***************************************************************************************************\n");

                    if (!bConnected || (null == service) || (null == service.Url))
                    {
                        Log.WriteLine("Failed to connect to exchange server for user " + sUserSmtp);
                        Console.WriteLine("Failed to connect to exchange server for user " + sUserSmtp);
                        break;
                    }

                    String AutoDiscoverURL = service.Url.ToString();
                    AutoDiscoverURL = AutoDiscoverURL.Trim();

                    if (String.Empty == AutoDiscoverURL)
                    {
                        Console.WriteLine("Failed to get exchange server for user " + sUserSmtp);
                        Log.WriteLine("Failed to get exchange server for user " + sUserSmtp);
                        break;
                    }

                    
                    //sFoldEwsId = "AAMkAGFkZTM1MjY3LWZiYzAtNDA1ZC04NWI3LTA1ZWRlYzE2NjVjZAAuAAAAAAAehyvl2c+VRaNBUFlASUlpAQA5Thqx2ogYS5z4GmODBiBuAAHgpLSfAAA="; //Deleted folder
                    FolderId id = new FolderId(sFoldEwsId);

                    Folder fld = Folder.Bind(service, id);
                    Log.WriteLine("");
                    Log.WriteLine("Folder Name: {0} - {1}", fld.DisplayName, fld.Id.UniqueId);
                    Log.WriteLine("");
                    Log.WriteLine("\n------------------------------------------------------------------------------------------------------\n");
                    sOutlookFolderName = fld.DisplayName;
                    if (fld.DisplayName.Length > 0)
                    {
                        //var view = new ItemView(100) { PropertySet = new PropertySet { EmailMessageSchema.Id, ItemSchema.Subject, ItemSchema.Id } };
                        //view.Traversal = ItemTraversal.Shallow;

                        String sDt = sStartDate;
                        if (sDt == "")
                            sDt = "2015-03-05";
                        sDt += "T23:59:50Z";
                        //SearchFilter.IsLessThan filter = new SearchFilter.IsLessThan(EmailMessageSchema.DateTimeReceived, sDt);//"2015-03-16T14:15:50Z");

                        ExtendedPropertyDefinition emailGuidProp = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders,
                                                                                                    "x-autn-guid", MapiPropertyType.String);

                        ExtendedPropertyDefinition filingStatus = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                                    "FilingStatus", MapiPropertyType.String);

                        SearchFilter.SearchFilterCollection searchFilterCollection =
                                                        new SearchFilter.SearchFilterCollection(LogicalOperator.And);

                        // Uncomment below line
                        searchFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.WorkSite.Ems.Queued"));
                        //searchFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.WorkSite.Ems.Filed"));

                        searchFilterCollection.Add(new SearchFilter.IsLessThan(EmailMessageSchema.DateTimeReceived, sDt));

                        FindItemsResults<Item> findResults;
                        //Collection<EmailMessage> 
                        
                        ItemView view = new ItemView(50,0,OffsetBasePoint.Beginning);

                        // Identify the Subject properties to return.
                        // Indicate that the base property will be the item identifier
                        view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, 
                                            emailGuidProp,                                                 
                                            filingStatus
                                            );

                        // Set the traversal to shallow. (Shallow is the default option; other options are Associated and SoftDeleted.)
                        view.Traversal = ItemTraversal.Shallow;

                        do
                        {

                            // Send the request to search the Inbox and get the results.
                            findResults = service.FindItems(id, searchFilterCollection, view);


                            int extendedPropertyindex = 0;
                            //bool bUpdate = false;

                            // Process each item.
                            foreach (Item myItem in findResults.Items)
                            {
                                extendedPropertyindex = 0;

                                if (myItem is EmailMessage)
                                {
                                    // Get EntryId from EWSId

                                    AlternateId objAltID = new AlternateId();
                                    objAltID.Format = IdFormat.EwsId;
                                    objAltID.Mailbox = sUserSmtp;
                                    objAltID.UniqueId = myItem.Id.ToString();

                                    //Convert  PR_ENTRYID identifier format to an EWS identifier. 
                                    AlternateIdBase objAltIDBase = service.ConvertId(objAltID, IdFormat.HexEntryId);
                                    AlternateId objAltIDResp = (AlternateId)objAltIDBase;

                                    // Check if this EntryId present in the EM_REQUEST
                                    if (m_oEntryId.ContainsKey(objAltIDResp.UniqueId))
                                    {
                                        Log.WriteLine("Skip - EntryId exist in EM_REQUEST - {0}", objAltIDResp.UniqueId);
                                        iSkippedQueuedEMails++;
                                        continue;
                                    }
                                    else
                                    {
                                        foreach (ExtendedProperty extendedProperty in myItem.ExtendedProperties)
                                        {
                                            if (extendedProperty.PropertyDefinition.Name == filingStatus.Name &&
                                                    extendedProperty.PropertyDefinition.PropertySetId == filingStatus.PropertySetId)
                                            {
                                                myItem.RemoveExtendedProperty(filingStatus);
                                                break;
                                            }

                                            extendedPropertyindex++;
                                        }

                                        foreach (ExtendedProperty extendedProperty in myItem.ExtendedProperties)
                                        {
                                            if (extendedProperty.PropertyDefinition.Name == emailGuidProp.Name &&
                                                    extendedProperty.PropertyDefinition.PropertySetId == emailGuidProp.PropertySetId)
                                            {
                                                myItem.RemoveExtendedProperty(emailGuidProp);
                                                break;
                                            }
                                           
                                            extendedPropertyindex++;
                                        }
                                        

                                        myItem.ItemClass = "IPM.Note";
                                        //bUpdate = true;

                                        if (sReportMode.ToUpper() == "FALSE")
                                            myItem.Update(ConflictResolutionMode.AlwaysOverwrite);

                                        Log.WriteLine("");
                                        Log.WriteLine("Updated : Subject - {0}", (myItem as EmailMessage).Subject);
                                        Log.WriteLine("Updated : EWSId - {0}", myItem.Id.UniqueId);
                                        Log.WriteLine("");
                                        Console.WriteLine("Updated : {0}", (myItem as EmailMessage).Subject);
                                        iTotalEmailsReset++;
                                    }
                                }
                            }

                            //if (bUpdate)
                            //{                                
                            //    if (sReportMode.ToUpper() == "FALSE")
                            //        service.UpdateItems(findResults, id, ConflictResolutionMode.AlwaysOverwrite, MessageDisposition.SaveOnly, null);                                
                                
                            //}
                            view.Offset += 50;
                        } while (findResults.MoreAvailable);

                    }
                    bRet = true;
                } while (false);
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("The specified object was not found in the store"))
                    Log.WriteLine(ex.Message);
                else
                    Log.WriteLine(ex.Message + ex.StackTrace);

                bRet = false;
            }
            Log.WriteLine("\n------------------------------------------------------------------------------------------------\n");
            Log.WriteLine("");
            Log.WriteLine("Total emails reset for {0} on Folder: {1} - {2} - {3}", sUserSmtp, sOutlookFolderName, sFolderEwsId, iTotalEmailsReset);
            Log.WriteLine("");
            Log.WriteLine("Total emails skipped for {0} on Folder: {1} - {2} - {3}", sUserSmtp, sOutlookFolderName, sFolderEwsId, iSkippedQueuedEMails);
            Log.WriteLine("");

            return bRet;
        }

        private bool ConnectToExchangeServer(String sExchangeSrv,
                                                String sImpersonatorSMTP, 
                                                String sImpersonatorPwd, 
                                                String sUserSMTP, 
                                                ref ExchangeService service)
        {
            bool bRet = false;
            

            try
            {
                service = new ExchangeService(this.ExchangeServerVer);

                service.Credentials = new WebCredentials(sImpersonatorSMTP, sImpersonatorPwd);
                service.TraceListener = new TraceListener();
                service.TraceFlags = TraceFlags.All;
                service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, sUserSMTP);
                service.TraceEnabled = true;
                ServicePointManager.ServerCertificateValidationCallback = Program.CertificateValidationCallback;

                if (String.Empty != sExchangeSrv)
                {
                    string exchangeUrl;
                    exchangeUrl = "https://";
                    exchangeUrl += sExchangeSrv;
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
                Log.WriteLine("Failed to connect to the exchange server. " + ex.Message);
                bRet = false;
            }

         
            return bRet;
        }

        public void ScanLinkedFoldersEx(string[] args)
        {

            if (args.Length < 10)
            {
                Console.WriteLine("Syntax: <Command> <ImpersonatorSMTP> <password> <end user> <exchange server name> <csvFile> <QueuedMsg> <Start Date> <End Date> <ReportMode>");
                Console.WriteLine("Example: SCAN-LINKED-FOLDERS ImpersonatorSMTPAddress@dev.local password user1@dev.local exchangeServer User1.csv True 2014-01-01 2016-02-22 True");
                return;
            }

            int iTotalCnt = 0;
            int iSkippedQueuedEMails = 0;
            long iTotalEmailsReset = 0;
            do
            {
                Log.AutoFlush = true;

                m_oEntryId = new Dictionary<String, String>();

                //if (!LoadEMProjectsEx(args[3], "WO-33282-2.csv", 2)) // 2- extract rows which has valid EntryId
                if (!LoadEMProjectsEx(args[3], "EmRequests.csv", 2)) // 2- extract rows which has valid EntryId
                {
                    Console.WriteLine(String.Format("Failed to load file EmRequests.csv"));
                    Log.WriteLine(String.Format("Failed to load file EmRequests.csv"));
                    break;
                }

                Dictionary<String, String> oFolderEntryIds = null;
                oFolderEntryIds = new Dictionary<String, String>();

                //StreamWriter Log = new StreamWriter("ScanLinkedFoldersOutput.txt", true);
                //Log.AutoFlush = true;

                // Create the binding.
                ExchangeService service;
                service = new ExchangeService(ExchangeVersion.Exchange2010);


                service.Credentials = new WebCredentials(args[1], args[2]);
                service.TraceListener = new TraceListener();
                service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

                string smtpAddress = args[3];
                service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

                String sExchServer = args[4];
                String sExchSrv;
                string[] exchArr = sExchServer.Split('>');
                if (exchArr.Count() > 1)
                    sExchSrv = exchArr[1];
                else if (sExchServer.Length > 0)
                    sExchSrv = sExchServer;
                else
                {
                    Log.WriteLine("Exchange server field is blank");
                    break;
                }

                string exchangeUrl;
                exchangeUrl = "https://";
                exchangeUrl += sExchSrv;//args[4];
                exchangeUrl += "/EWS/Exchange.asmx";


                service.Url = new Uri(exchangeUrl);


                Console.WriteLine("AutodiscoverURL: " + service.Url);

                ServicePointManager.ServerCertificateValidationCallback = Program.CertificateValidationCallback;

                service.TraceEnabled = true;

                String CSVFilePath = args[5];
                if (!File.Exists(CSVFilePath))
                {
                    Log.WriteLine("File doesn't exist - {0}", CSVFilePath);
                    Console.WriteLine("File doesn't exist - {0}", CSVFilePath);
                    break;
                }

                System.IO.StreamReader file = new System.IO.StreamReader(CSVFilePath);

                string line;
                while ((line = file.ReadLine()) != null)
                {
                    line.Trim();
                    String[] Tokens = line.Split(",".ToCharArray());
                    if (2 > Tokens.Length)
                    {
                        Log.WriteLine("Invalid entry in {0}", CSVFilePath);
                        break;
                    }
                    if (!oFolderEntryIds.ContainsKey(Tokens[1].ToString()))
                        oFolderEntryIds.Add(Tokens[1].ToString(), Tokens[0].ToString());
                    else
                        Console.WriteLine("Record Exist");
                }

                foreach (KeyValuePair<String, String> Entry in oFolderEntryIds)
                {
                    try
                    {
                        ////////////
                        AlternateId oAltID = new AlternateId();
                        oAltID.Format = IdFormat.HexEntryId;
                        oAltID.Mailbox = smtpAddress;
                        oAltID.UniqueId = Entry.Value;

                        //Convert  PR_ENTRYID identifier format to an EWS identifier. 
                        AlternateIdBase oAltIDBase = service.ConvertId(oAltID, IdFormat.EwsId);
                        AlternateId oAltIDResp = (AlternateId)oAltIDBase;

                        ////////////
                        String FoldEwsId = oAltIDResp.UniqueId; //Entry.Key;
                        String FoldName = Entry.Value;
                        int iSkippedForThisFolder = 0;
                        int iCount = 0;


                        Folder fld;
                        FolderId id = new FolderId(FoldEwsId);

                        fld = Folder.Bind(service, id);
                        Console.WriteLine("Folder Name: " + fld.DisplayName);
                        FoldName = fld.DisplayName;

                        SearchFilter.SearchFilterCollection searchAndFilterCollection =
                                                new SearchFilter.SearchFilterCollection(LogicalOperator.And);

                        SearchFilter.SearchFilterCollection searchOrFilterCollection =
                                                new SearchFilter.SearchFilterCollection(LogicalOperator.Or);

                        ExtendedPropertyDefinition emailGuidProp = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders,
                                                                                               "x-autn-guid", MapiPropertyType.String);

                        ExtendedPropertyDefinition filingStatus = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                                    "FilingStatus", MapiPropertyType.String);


                        ExtendedPropertyDefinition lastChangeDt = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                    "x-autn-lastchange-time", MapiPropertyType.SystemTime);


                        String sQueueOnly = args[6];
                        String sStartDt = args[7];
                        String sEndDt = args[8];
                        //if (sStartDt == "")
                        //    sStartDt = "2014-10-01";
                        //if (sEndDt == "")
                        //    sEndDt = "2015-07-01";
                        
                        String sReportMode = args[9];

                        //searchFilterCollection.Add(new SearchFilter.IsGreaterThan(lastChangeDt, sStartDt));
                        //searchFilterCollection.Add(new SearchFilter.IsGreaterThan(EmailMessageSchema.DateTimeReceived, sStartDt));

                        //searchFilterCollection.Add(new SearchFilter.IsLessThan(lastChangeDt, sEndDt));
                        //searchFilterCollection.Add(new SearchFilter.IsLessThan(EmailMessageSchema.DateTimeReceived, sEndDt));

                        if (!File.Exists("MessageClasses.txt"))
                        {
                            Log.WriteLine("File doesn't exist - MessageClasses.txt");
                            Console.WriteLine("File doesn't exist - MessageClasses.txt");
                            break;
                        }

                        System.IO.StreamReader fileMsgClass = new System.IO.StreamReader("MessageClasses.txt");

                        string lineMsgCls;
                        while ((lineMsgCls = fileMsgClass.ReadLine()) != null)
                        {
                            lineMsgCls.Trim();
                            if (lineMsgCls.Length > 1)
                                searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, lineMsgCls));
                        }

                        //searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.WorkSite.Ems.Queued"));//args[6]));//"IPM.Note.WorkSite.Ems.Filed"));
                        //searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.ABCD"));
                        if (sQueueOnly.ToUpper() == "TRUE")
                            searchAndFilterCollection.Add(new SearchFilter.IsEqualTo(filingStatus, "Queued"));                         
                        else
                        { 
                            searchAndFilterCollection.Add(new SearchFilter.IsEqualTo(filingStatus, "Filed"));  
                            searchAndFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.WorkSite.Ems.Queued"));                              
                        }

                        //searchAndFilterCollection.Add(new SearchFilter.IsEqualTo(filingStatus, "Filed"));                         
                        searchAndFilterCollection.Add(new SearchFilter.Exists(emailGuidProp));

                        if (sStartDt.Length > 1)
                        {
                            sStartDt += "T00:00:00Z";
                            searchAndFilterCollection.Add(new SearchFilter.IsGreaterThan(lastChangeDt, sStartDt));
                        }
                        if (sEndDt.Length > 1)
                        {
                            sEndDt += "T23:59:50Z";
                            searchAndFilterCollection.Add(new SearchFilter.IsLessThan(lastChangeDt, sEndDt));
                        }
                        
                        
                        searchAndFilterCollection.Add(searchOrFilterCollection);

                        FindItemsResults<Item> findResults;
                        ItemView view = new ItemView(100, 0, OffsetBasePoint.Beginning);

                        // Identify the Subject properties to return.
                        // Indicate that the base property will be the item identifier
                        view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.ItemClass, emailGuidProp, filingStatus);

                        // Set the traversal to shallow. (Shallow is the default option; other options are Associated and SoftDeleted.)
                        view.Traversal = ItemTraversal.Shallow;



                        do
                        {


                            // Send the request to search the Inbox and get the results.
                            findResults = service.FindItems(id, searchAndFilterCollection, view);
                            //findResults = service.FindItems(id, view);

                            bool bUpdate = true;

                            if (bUpdate)
                            {
                                int extendedPropertyindex = 0;
                                //bool bUpdate = false;

                                // Process each item.
                                foreach (Item myItem in findResults.Items)
                                {
                                    extendedPropertyindex = 0;

                                    if (myItem is EmailMessage)
                                    {
                                        // Get EntryId from EWSId

                                        AlternateId objAltID = new AlternateId();
                                        objAltID.Format = IdFormat.EwsId;
                                        objAltID.Mailbox = smtpAddress;
                                        objAltID.UniqueId = myItem.Id.ToString();

                                        //Convert  PR_ENTRYID identifier format to an EWS identifier. 
                                        AlternateIdBase objAltIDBase = service.ConvertId(objAltID, IdFormat.HexEntryId);
                                        AlternateId objAltIDResp = (AlternateId)objAltIDBase;

                                        // Check if this EntryId present in the EM_REQUEST
                                        if (m_oEntryId.ContainsKey(objAltIDResp.UniqueId))
                                        {
                                            Log.WriteLine("Skip - EntryId {0} exist in EM_REQUEST", objAltIDResp.UniqueId);
                                            iSkippedQueuedEMails++;
                                            iSkippedForThisFolder++;
                                            continue;
                                        }
                                        else
                                        {
                                            foreach (ExtendedProperty extendedProperty in myItem.ExtendedProperties)
                                            {
                                                if (extendedProperty.PropertyDefinition.Name == filingStatus.Name &&
                                                        extendedProperty.PropertyDefinition.PropertySetId == filingStatus.PropertySetId)
                                                {
                                                    myItem.RemoveExtendedProperty(filingStatus);
                                                    break;
                                                }

                                                extendedPropertyindex++;
                                            }

                                            foreach (ExtendedProperty extendedProperty in myItem.ExtendedProperties)
                                            {
                                                if (extendedProperty.PropertyDefinition.Name == emailGuidProp.Name &&
                                                        extendedProperty.PropertyDefinition.PropertySetId == emailGuidProp.PropertySetId)
                                                {
                                                    myItem.RemoveExtendedProperty(emailGuidProp);
                                                    break;
                                                }

                                                extendedPropertyindex++;
                                            }


                                            if ((myItem.ItemClass.ToUpper() == "IPM.NOTE.WORKSITE.EMS.QUEUED") ||
                                                (myItem.ItemClass.ToUpper() == "IPM.NOTE.WORKSITE.EMS.FILED"))
                                                myItem.ItemClass = "IPM.Note";
                                                
                                            
                                            bUpdate = true;

                                            if (sReportMode.ToUpper() == "FALSE")
                                                myItem.Update(ConflictResolutionMode.AlwaysOverwrite);

                                            Log.WriteLine("Updated : Subject - {0}", (myItem as EmailMessage).Subject);
                                            Log.WriteLine("Updated : EWSId - {0}", myItem.Id.UniqueId);
                                            Log.WriteLine("");
                                            Console.WriteLine("Updated : {0}", (myItem as EmailMessage).Subject);
                                            iTotalEmailsReset++;
                                        }
                                    }
                                }
                            }

                            // Process each item.
                            //foreach (Item myItem in findResults.Items)
                            //{
                            //    if (myItem is EmailMessage)
                            //    {
                            //        iCount++;
                            //        Console.WriteLine((myItem as EmailMessage).Subject);
                            //    }
                            //}
                            iCount += findResults.Items.Count();
                            iTotalCnt += findResults.Items.Count();
                            if (sReportMode.ToUpper() == "TRUE")
                                view.Offset += 100;
                            else
                                view.Offset = iSkippedForThisFolder;
                        } while (findResults.MoreAvailable);


                        Console.WriteLine("Reset count : {0} ", iTotalEmailsReset);
                        //Log.WriteLine("Folder : {0} : Items Processed : {1}", FoldName, iTotalCnt);

                        Log.WriteLine("");
                        Log.WriteLine("Total emails reset for {0} on Folder: {1} - {2} - {3}", smtpAddress, FoldName, FoldEwsId, iTotalEmailsReset);
                        Log.WriteLine("");
                        Log.WriteLine("Total emails skipped (request exist in em_req) for {0} on Folder: {1} - {2} - {3}", smtpAddress, FoldName, FoldEwsId, iSkippedQueuedEMails);
                        Log.WriteLine("");

                    }
                    catch (Exception ex)
                    {
                        Log.WriteLine("Folder: {0} : ", ex.Message);
                    }
                }
                Log.WriteLine("Total Items Processed : {0} ", iTotalCnt);
            } while (false);


        }

        public void ScanLinkedFolders(string[] args)
        {

            if (args.Length < 9)
            {
                Console.WriteLine("Syntax: <Command> <ImpersonatorSMTP> <password> <end user> <exchange server name> <csvFile> <MsgClass> <Start Date> <End Date>");
                Console.WriteLine("Example: SCAN-LINKED-FOLDERS ImpersonatorSMTPAddress@dev.local password user1@dev.local xchange.dev.local LinkedFolderList.csv IPM.Note 2015-06-19 2015-06-20");
                return;
            }

            int iTotalCnt = 0;
            long iSkippedQueuedEMails = 0;
            long iTotalEmailsReset = 0;
            do
            {
                Log.AutoFlush = true;

                m_oEntryId = new Dictionary<String, String>();

                if (!LoadEMProjectsEx(args[3], "WO-33282-2.csv", 2)) // 2- extract rows which has valid EntryId
                {
                    Console.WriteLine(String.Format("Failed to load file WO-33282-2.csv"));
                    Log.WriteLine(String.Format("Failed to load file WO-33282-2.csv"));
                    break;
                }

                Dictionary<String, String> oFolderEntryIds = null;
                oFolderEntryIds = new Dictionary<String, String>();

                //StreamWriter Log = new StreamWriter("ScanLinkedFoldersOutput.txt", true);
                //Log.AutoFlush = true;

                // Create the binding.
                ExchangeService service;
                service = new ExchangeService(ExchangeVersion.Exchange2010);


                service.Credentials = new WebCredentials(args[1], args[2]);
                service.TraceListener = new TraceListener();
                service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

                string smtpAddress = args[3];
                service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

                String sExchServer = args[4];
                String sExchSrv;
                string[] exchArr = sExchServer.Split('>');
                if (exchArr.Count() > 1)
                    sExchSrv = exchArr[1];
                else if (sExchServer.Length > 0)
                    sExchSrv = sExchServer;
                else
                {
                    Log.WriteLine("Exchange server field is blank");
                    break;
                }

                string exchangeUrl;
                exchangeUrl = "https://";
                exchangeUrl += sExchSrv;//args[4];
                exchangeUrl += "/EWS/Exchange.asmx";


                service.Url = new Uri(exchangeUrl);


                Console.WriteLine("AutodiscoverURL: " + service.Url);

                ServicePointManager.ServerCertificateValidationCallback = Program.CertificateValidationCallback;

                service.TraceEnabled = true;

                String CSVFilePath = args[5];
                if (!File.Exists(CSVFilePath))
                {
                    Log.WriteLine("File doesn't exist - {0}", CSVFilePath);
                    break;
                }

                System.IO.StreamReader file = new System.IO.StreamReader(CSVFilePath);

                string line;
                while ((line = file.ReadLine()) != null)
                {
                    line.Trim();
                    String[] Tokens = line.Split(",".ToCharArray());
                    if (2 > Tokens.Length)
                    {
                        Log.WriteLine("Invalid entry in {0}", CSVFilePath);
                        break;
                    }
                    if (!oFolderEntryIds.ContainsKey(Tokens[1].ToString()))
                        oFolderEntryIds.Add(Tokens[1].ToString(), Tokens[0].ToString());
                    else
                        Console.WriteLine("Record Exist");
                }

                foreach (KeyValuePair<String, String> Entry in oFolderEntryIds)
                {
                    try
                    {
                        String FoldEwsId = Entry.Key;
                        String FoldName = Entry.Value;
                        int iCount = 0;


                        Folder fld;
                        FolderId id = new FolderId(FoldEwsId);

                        fld = Folder.Bind(service, id);
                        Console.WriteLine("Folder Name: " + fld.DisplayName);
                        FoldName = fld.DisplayName;

                        SearchFilter.SearchFilterCollection searchFilterCollection =
                                                new SearchFilter.SearchFilterCollection(LogicalOperator.And);


                        ExtendedPropertyDefinition emailGuidProp = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders,
                                                                                               "x-autn-guid", MapiPropertyType.String);

                        ExtendedPropertyDefinition filingStatus = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                                    "FilingStatus", MapiPropertyType.String);


                        ExtendedPropertyDefinition lastChangeDt = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                    "x-autn-lastchange-time", MapiPropertyType.SystemTime);


                        String sStartDt = "2014-08-01";//args[7];
                        String sEndDt = args[7];//args[8];
                        if (sStartDt == "")
                            sStartDt = "2014-10-01";
                        if (sEndDt == "")
                            sEndDt = "2015-07-01";
                        sStartDt += "T00:00:00Z";
                        sEndDt += "T23:59:50Z";
                        String sReportMode = args[8];//args[9];
                        Log.WriteLine("SD - {0}, ED - {1}", sStartDt, sEndDt); 
                        //searchFilterCollection.Add(new SearchFilter.IsGreaterThan(lastChangeDt, sStartDt));
                        searchFilterCollection.Add(new SearchFilter.IsGreaterThan(EmailMessageSchema.DateTimeReceived, sStartDt));
                        
                        //searchFilterCollection.Add(new SearchFilter.IsLessThan(lastChangeDt, sEndDt));
                        searchFilterCollection.Add(new SearchFilter.IsLessThan(EmailMessageSchema.DateTimeReceived, sEndDt));

                        searchFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note"));//args[6]));//"IPM.Note.WorkSite.Ems.Filed"));
                        Log.WriteLine("IPM.Note11111");
                        FindItemsResults<Item> findResults;
                        ItemView view = new ItemView(100, 0, OffsetBasePoint.Beginning);

                        // Identify the Subject properties to return.
                        // Indicate that the base property will be the item identifier
                        view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, emailGuidProp,filingStatus);

                        // Set the traversal to shallow. (Shallow is the default option; other options are Associated and SoftDeleted.)
                        view.Traversal = ItemTraversal.Shallow;



                        do
                        {


                            // Send the request to search the Inbox and get the results.
                            findResults = service.FindItems(id, searchFilterCollection, view);
                            //findResults = service.FindItems(id, view);

                            bool bUpdate = true;

                            if (bUpdate)
                            {
                                int extendedPropertyindex = 0;
                                //bool bUpdate = false;

                                // Process each item.
                                foreach (Item myItem in findResults.Items)
                                {
                                    extendedPropertyindex = 0;

                                    if (myItem is EmailMessage)
                                    {
                                        // Get EntryId from EWSId

                                        AlternateId objAltID = new AlternateId();
                                        objAltID.Format = IdFormat.EwsId;
                                        objAltID.Mailbox = smtpAddress;
                                        objAltID.UniqueId = myItem.Id.ToString();

                                        //Convert  PR_ENTRYID identifier format to an EWS identifier. 
                                        AlternateIdBase objAltIDBase = service.ConvertId(objAltID, IdFormat.HexEntryId);
                                        AlternateId objAltIDResp = (AlternateId)objAltIDBase;

                                        // Check if this EntryId present in the EM_REQUEST
                                        if (m_oEntryId.ContainsKey(objAltIDResp.UniqueId))
                                        {
                                            Log.WriteLine("Skip - EntryId {0} exist in EM_REQUEST", objAltIDResp.UniqueId);
                                            iSkippedQueuedEMails++;
                                            continue;
                                        }
                                        else
                                        {
                                            foreach (ExtendedProperty extendedProperty in myItem.ExtendedProperties)
                                            {
                                                if (extendedProperty.PropertyDefinition.Name == filingStatus.Name &&
                                                        extendedProperty.PropertyDefinition.PropertySetId == filingStatus.PropertySetId)
                                                {
                                                    myItem.RemoveExtendedProperty(filingStatus);
                                                    break;
                                                }

                                                extendedPropertyindex++;
                                            }

                                            foreach (ExtendedProperty extendedProperty in myItem.ExtendedProperties)
                                            {
                                                if (extendedProperty.PropertyDefinition.Name == emailGuidProp.Name &&
                                                        extendedProperty.PropertyDefinition.PropertySetId == emailGuidProp.PropertySetId)
                                                {
                                                    myItem.RemoveExtendedProperty(emailGuidProp);
                                                    break;
                                                }

                                                extendedPropertyindex++;
                                            }


                                            myItem.ItemClass = "IPM.Note";
                                            bUpdate = true;

                                            if (sReportMode.ToUpper() == "FALSE")
                                                myItem.Update(ConflictResolutionMode.AlwaysOverwrite);

                                            Log.WriteLine("Updated : Subject - {0}", (myItem as EmailMessage).Subject);
                                            Log.WriteLine("Updated : EWSId - {0}", myItem.Id.UniqueId);
                                            Log.WriteLine("");
                                            Console.WriteLine("Updated : {0}", (myItem as EmailMessage).Subject);
                                            iTotalEmailsReset++;                                            
                                        }
                                    }
                                }
                            }

                            // Process each item.
                            //foreach (Item myItem in findResults.Items)
                            //{
                            //    if (myItem is EmailMessage)
                            //    {
                            //        iCount++;
                            //        Console.WriteLine((myItem as EmailMessage).Subject);
                            //    }
                            //}
                            iCount += findResults.Items.Count();
                            iTotalCnt += findResults.Items.Count();
                            view.Offset += 100;
                        } while (findResults.MoreAvailable);


                        Console.WriteLine("Reset count : {0} ", iTotalEmailsReset);
                        //Log.WriteLine("Folder : {0} : Items Processed : {1}", FoldName, iTotalCnt);

                        Log.WriteLine("");
                        Log.WriteLine("Total emails reset for {0} on Folder: {1} - {2} - {3}", smtpAddress, FoldName, FoldEwsId, iTotalEmailsReset);
                        Log.WriteLine("");
                        Log.WriteLine("Total emails skipped (request exist in em_req) for {0} on Folder: {1} - {2} - {3}", smtpAddress, FoldName, FoldEwsId, iSkippedQueuedEMails);
                        Log.WriteLine("");

                    }
                    catch (Exception ex)
                    {
                        Log.WriteLine("Folder: {0} : ", ex.Message);
                    }
                }
                Log.WriteLine("Total Items Processed : {0} ", iTotalCnt);
            } while (false);


        }

        private bool LoadEMProjectsEx(String sInputUser, String CSVFilePath, int iOperation)
        {
            bool bRet = false;
            do
            {
                if (!File.Exists(CSVFilePath))
                {
                    Log.WriteLine("File doesn't exist - {0}", CSVFilePath);
                    break;
                }

                System.IO.StreamReader file = new System.IO.StreamReader(CSVFilePath);
                string line;
                //bool bProcess = false;
                Int32 sLineNum = 0;
                String sEntryId;
                String sFolderPath;
                String sUserSMTP;
                String sServer;
                String sEmailGuid;
                bool bInvalidUserDisplayed = false;

                while ((line = file.ReadLine()) != null)
                {
                    line.Trim();
                    if (String.IsNullOrEmpty(line))
                    {
                        Log.WriteLine("line is empty - {0} - {1}", sInputUser, CSVFilePath);
                        continue;
                    }

                    String[] Tokens = line.Split(",".ToCharArray());
                    if (5 > Tokens.Length)
                    {
                        throw new Exception(String.Format("Invalid entry in {0} at line {1}", CSVFilePath, sLineNum));
                    }

                    sEntryId = Tokens[0].ToUpper(); // EM_REQUEST - MSG_ID
                    sEmailGuid = Tokens[1].ToUpper(); // EM_REQUEST - EMAIL_GUID
                    sFolderPath = Tokens[2]; // EM_REQUEST - FOLDER_PATH
                    sUserSMTP = Tokens[3].ToUpper(); // DOCUSER - EMAIL
                    sServer = Tokens[4].ToUpper(); // DOCUSER - EXCH_AUTO_DISC

                    if (sEntryId == "NULL")
                        sEntryId = "";
                    if (sEmailGuid == "NULL")
                        sEmailGuid = "";
                    if ((sFolderPath == "null") || (sFolderPath == "NULL"))
                        sFolderPath = "";
                    if (sUserSMTP == "NULL")
                        sUserSMTP = "";
                    if (sServer == "NULL")
                        sServer = "";


                    // Process only the user specified in the command ilne
                    if (sUserSMTP != sInputUser.ToUpper())
                    {
                        //if (!bInvalidUserDisplayed)
                        //{
                        //    Log.WriteLine("User provided in Command line and user in CSV file are not matching - {0} - {1} - {2} ", sUserSMTP, sInputUser, CSVFilePath);
                        //    Console.WriteLine("User provided in Command line and user in CSV file are not matching - {0} - {1} - {2} ", sUserSMTP, sInputUser, CSVFilePath);
                        //    bInvalidUserDisplayed = true;
                        //}
                        continue;
                    }

                    if (sEntryId.Length > 0)
                    {
                        if (!m_oEntryId.ContainsKey(sEntryId))
                            m_oEntryId.Add(sEntryId, sUserSMTP);
                    }                   
                    
                }
                bRet = true;
            } while (false);


            return bRet;
        }

        private bool LoadEMProjects(String sInputUser, String CSVFilePath, int iOperation)
        {
            bool bRet = false;
            do
            {
                if (!File.Exists(CSVFilePath))
                {
                    Log.WriteLine("File doesn't exist - {0}", CSVFilePath);
                    break;
                }

                System.IO.StreamReader file = new System.IO.StreamReader(CSVFilePath);
                string line;
                //bool bProcess = false;
                Int32 sLineNum = 0;
                String sEntryId;
                String sFolderPath;
                String sUserSMTP;
                String sServer;
                String sEmailGuid;
                bool bInvalidUserDisplayed = false;

                while ((line = file.ReadLine()) != null)
                {
                    line.Trim();
                    if (String.IsNullOrEmpty(line))
                    {
                        Log.WriteLine("line is empty - {0} - {1}", sInputUser, CSVFilePath);
                        continue;
                    }

                    String[] Tokens = line.Split(",".ToCharArray());
                    if (5 > Tokens.Length)
                    {
                        throw new Exception(String.Format("Invalid entry in {0} at line {1}", CSVFilePath, sLineNum));
                    }

                    sEntryId = Tokens[0].ToUpper(); // EM_REQUEST - MSG_ID
                    sEmailGuid = Tokens[1].ToUpper(); // EM_REQUEST - EMAIL_GUID
                    sFolderPath = Tokens[2]; // EM_REQUEST - FOLDER_PATH
                    sUserSMTP = Tokens[3].ToUpper(); // DOCUSER - EMAIL
                    sServer = Tokens[4].ToUpper(); // DOCUSER - EXCH_AUTO_DISC

                    if (sEntryId == "NULL")
                        sEntryId = "";
                    if (sEmailGuid == "NULL")
                        sEmailGuid = "";
                    if ((sFolderPath == "null") || (sFolderPath == "NULL"))
                        sFolderPath = "";
                    if (sUserSMTP == "NULL")
                        sUserSMTP = "";
                    if (sServer == "NULL")
                        sServer = "";

                   
                    // Process only the user specified in the command ilne
                    if (sUserSMTP != sInputUser.ToUpper())
                    {
                        if (!bInvalidUserDisplayed)
                        {
                            Log.WriteLine("User provided in Command line and user in CSV file are not matching - {0} - {1} - {2} ", sUserSMTP, sInputUser, CSVFilePath);
                            Console.WriteLine("User provided in Command line and user in CSV file are not matching - {0} - {1} - {2} ", sUserSMTP, sInputUser, CSVFilePath);
                            bInvalidUserDisplayed = true;
                        }
                        continue;
                    }

                    if (sEntryId.Length > 0)
                    {
                        if (!m_oEntryId.ContainsKey(sEntryId))
                            m_oEntryId.Add(sEntryId, sUserSMTP);
                    }

                    bool bValidate = false;
                    if (iOperation == 1)
                    {
                        // Process only NULL EntryId and NULL EmailGuid rows
                        if ((sEntryId.Length <= 0) && (sEmailGuid.Length <= 1))
                            bValidate = true;
                    }
                    else if (iOperation == 2)
                    {
                        if (sEntryId.Length > 1)
                            bValidate = true;
                    }
                    
                    if (bValidate)
                    {                        
                        int nIndex1 = sUserSMTP.IndexOf('@');
                        if (1 > nIndex1)
                        {
                            Log.WriteLine("Invalid email address - {0} - {1}", sUserSMTP, CSVFilePath); 
                            continue;
                        }

                        int nIndex2 = sUserSMTP.Substring(nIndex1 + 1).IndexOf('.');
                        if (1 > nIndex2)
                        {
                            Log.WriteLine("Invalid email address- {0} - {1}", sUserSMTP, CSVFilePath); 
                            continue;
                        }

                        if (!m_oMarkedFolders.ContainsKey(sUserSMTP))
                        {
                            Dictionary<String, String> UserMarkedFolder = new Dictionary<String, String>();
                            
                            if (sFolderPath.Length > 0)
                            {                            
                                UserMarkedFolder.Add(sFolderPath, sServer);

                                m_oMarkedFolders.Add(sUserSMTP, UserMarkedFolder);
                            }
                            
                        }
                        else
                        {
                            Dictionary<String, String> UserMarkedFolders = m_oMarkedFolders[sUserSMTP];
                            try
                            {
                                if (!UserMarkedFolders.ContainsKey(sFolderPath))
                                {
                                    //Dictionary<String, String> UserMrkedFldr = new Dictionary<String, String>();
                                    UserMarkedFolders.Add(sFolderPath, sServer);                               
                                }
                            }
                            catch
                            {
                                Log.WriteLine(String.Format("Found duplicate folder entry id {0} for user {1}", sFolderPath, sUserSMTP));
                            }
                        }
                    }
                }
                bRet = true;
            } while (false);

           
            return bRet;
        }


        public void ExecuteScanOutlookFolders(string[] args)        
        {
            try
            {
                do
                {
                    Log.AutoFlush = true;

                    if (args.Length < 7)
                    {
                        Console.WriteLine("");
                        Console.WriteLine("FAILED!!! - Invalid parameters");
                        Console.WriteLine("");
                        Console.WriteLine("Syntax: <Command> <ImpersonatorSMTP> <password> <endUserSMTP> <exchange server name> <User EM_REQUEST CSV filePath> <RunReportMode>");
                        //SCAN-FOLDERS admin2@imanage.microsoftonline.com !wov2014 jsmith@imanage.microsoftonline.com ch1prd0410.outlook.com d:\Resubmit1.csv 2015-03-16
                        Console.WriteLine("Example: SCAN-OUTLOOK-FOLDERS ImpersonatorSMTPAddress@dev.local password endUserSMTPAddress xchange.dev.local c:\\User.csv True");
                        
                        break;
                    }

                    //m_oMarkedFolders = new Dictionary<String, List<String>>();
                    m_oEntryId = new Dictionary<String, String>();

                    if (!LoadEMProjects(args[3], args[5], 2)) // 2- extract rows which has valid EntryId
                    {
                        Console.WriteLine(String.Format("Failed to load file {0}", args[5]));
                        Log.WriteLine(String.Format("Failed to load file {0}", args[5]));
                        break;
                    }

                    Log.WriteLine("SCAN-OUTLOOK-FOLDERS");
                    foreach (KeyValuePair<String, Dictionary<String, String>> Entry in m_oMarkedFolders)
                    {
                        String UserSmtp = Entry.Key;
                        Dictionary<String, String> UserFolders = Entry.Value;

                        Log.WriteLine("\n");
                        Log.WriteLine(String.Format("Processing user {0}\n", UserSmtp));
                        Log.WriteLine("\n***************************************************************************************************\n");
                        foreach (KeyValuePair<String, String> Folder in UserFolders)
                        {
                            Log.WriteLine(String.Format("Folder EWSId - {0} : ", Folder.Key));
                            Log.WriteLine(String.Format("Exchange server - {0} ", Folder.Value));

                            if (!ScanAndUpdateOutlookFolder(args[1], args[2], UserSmtp, Folder.Key, Folder.Value, args[6], args[7]))
                                Log.WriteLine(String.Format("Failed to process Folder - EWSId - {0} : ", Folder.Key));
                            else
                                Log.WriteLine(String.Format("Processed Folder - EWSId - {0} : ", Folder.Key));
                            Log.WriteLine("\n");
                            Log.WriteLine("\n***************************************************************************************************\n");

                        }
                        
                        
                    }

                }
                while (false);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
            }
            finally
            {
                
            }
        }

        private bool ScanAndUpdateOutlookFolder(String sImpersonatorSMTP,
                                        String sImpersonatorPwd,
                                        String sUserSmtp,
                                        String sFolderEwsId,
                                        String sExchServer,
                                        String sStartDate,                            
                                        String sReportMode)
        {
            bool bRet = false;
            String sUserSmtpAddr;
            String sFoldEwsId;
            String sExchSrv;
            String sOutlookFolderName = "";
            long iTotalEmailsReset = 0;
            long iSkippedQueuedEMails = 0;
            try
            {
                do
                {
                    if ((sUserSmtp.Length == 0) || (sFolderEwsId.Length == 0) || (sExchServer.Length == 0))
                    {
                        Log.WriteLine("Invalid param {0} - {1} - {2}", sUserSmtp, sFolderEwsId, sExchServer);
                        break;
                    }

                    
                    // Get Folder EWS ID
                    sUserSmtpAddr = sUserSmtp;
                    string[] foldArr = sFolderEwsId.Split(':');
                    if (foldArr.Count() <= 1)
                    {
                        Log.WriteLine("Improper folder-id {0}", sFolderEwsId);
                        break;
                    }
                    sFoldEwsId = foldArr[1];

                    // Get Exchange server name
                    string[] exchArr = sExchServer.Split('>');
                    if (exchArr.Count() > 1)
                        sExchSrv = exchArr[1];
                    else if (sExchServer.Length > 0)
                        sExchSrv = sExchServer;
                    else
                    {
                        Log.WriteLine("Exchange server field is blank");
                        break;
                    }


                    ExchangeService service = null;
                    bool bConnected = ConnectToExchangeServer(sExchSrv, sImpersonatorSMTP, sImpersonatorPwd, sUserSmtp, ref service);

                    Log.WriteLine("\n***************************************************************************************************\n");

                    if (!bConnected || (null == service) || (null == service.Url))
                    {
                        Log.WriteLine("Failed to connect to exchange server for user " + sUserSmtp);
                        Console.WriteLine("Failed to connect to exchange server for user " + sUserSmtp);
                        break;
                    }

                    String AutoDiscoverURL = service.Url.ToString();
                    AutoDiscoverURL = AutoDiscoverURL.Trim();

                    if (String.Empty == AutoDiscoverURL)
                    {
                        Console.WriteLine("Failed to get exchange server for user " + sUserSmtp);
                        Log.WriteLine("Failed to get exchange server for user " + sUserSmtp);
                        break;
                    }

                    if (sFolderEwsId.Contains("EwsFolderId:"))
                    {
                        AlternateId objAltID = new AlternateId();
                        objAltID.Format = IdFormat.HexEntryId;
                        objAltID.Mailbox = sUserSmtp;
                        objAltID.UniqueId = sFolderEwsId.Substring(12);

                        //Convert  PR_ENTRYID identifier format to an EWS identifier. 
                        AlternateIdBase objAltIDBase = service.ConvertId(objAltID, IdFormat.EwsId);
                        AlternateId objAltIDResp = (AlternateId)objAltIDBase;
                        sFoldEwsId = objAltIDResp.UniqueId;
                    }

                    //sFoldEwsId = "AAMkAGFkZTM1MjY3LWZiYzAtNDA1ZC04NWI3LTA1ZWRlYzE2NjVjZAAuAAAAAAAehyvl2c+VRaNBUFlASUlpAQA5Thqx2ogYS5z4GmODBiBuAAHgpLSfAAA="; //Deleted folder
                    FolderId id = new FolderId(sFoldEwsId);

                    Folder fld = Folder.Bind(service, id);
                    Log.WriteLine("");
                    Log.WriteLine("Folder Name: {0} - {1}", fld.DisplayName, fld.Id.UniqueId);
                    Log.WriteLine("");
                    Log.WriteLine("\n------------------------------------------------------------------------------------------------------\n");
                    sOutlookFolderName = fld.DisplayName;
                    if (fld.DisplayName.Length > 0)
                    {
                        String sDt = sStartDate;
                        if (sDt == "")
                            sDt = "2015-03-05";
                        sDt += "T23:59:50Z";

                        ExtendedPropertyDefinition emailGuidProp = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders,
                                                                                                    "x-autn-guid", MapiPropertyType.String);

                        ExtendedPropertyDefinition filingStatus = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                                    "FilingStatus", MapiPropertyType.String);

                        SearchFilter.SearchFilterCollection searchFilterCollection =
                                                        new SearchFilter.SearchFilterCollection(LogicalOperator.And);

                        // Uncomment below line
                        //searchFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.WorkSite.Ems.Queued"));
                        searchFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.WorkSite.Ems.Queued"));

                        searchFilterCollection.Add(new SearchFilter.IsLessThan(EmailMessageSchema.DateTimeReceived, sDt));

                        FindItemsResults<Item> findResults;
                        //Collection<EmailMessage> 

                        ItemView view = new ItemView(50, 0, OffsetBasePoint.Beginning);

                        // Identify the Subject properties to return.
                        // Indicate that the base property will be the item identifier
                        view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject,
                                            emailGuidProp,
                                            filingStatus
                                            );

                        // Set the traversal to shallow. (Shallow is the default option; other options are Associated and SoftDeleted.)
                        view.Traversal = ItemTraversal.Shallow;

                        do
                        {

                            // Send the request to search the Inbox and get the results.
                            findResults = service.FindItems(id, searchFilterCollection, view);


                            int extendedPropertyindex = 0;
                            //bool bUpdate = false;

                            // Process each item.
                            foreach (Item myItem in findResults.Items)
                            {
                                extendedPropertyindex = 0;

                                if (myItem is EmailMessage)
                                {
                                    // Get EntryId from EWSId

                                    AlternateId objAltID = new AlternateId();
                                    objAltID.Format = IdFormat.EwsId;
                                    objAltID.Mailbox = sUserSmtp;
                                    objAltID.UniqueId = myItem.Id.ToString();

                                    //Convert  PR_ENTRYID identifier format to an EWS identifier. 
                                    AlternateIdBase objAltIDBase = service.ConvertId(objAltID, IdFormat.HexEntryId);
                                    AlternateId objAltIDResp = (AlternateId)objAltIDBase;

                                    // Check if this EntryId present in the EM_REQUEST
                                    if (m_oEntryId.ContainsKey(objAltIDResp.UniqueId))
                                    {
                                        Log.WriteLine("Skip - EntryId {0} exist in EM_REQUEST", objAltIDResp.UniqueId);
                                        iSkippedQueuedEMails++;
                                        continue;
                                    }
                                    else
                                    {
                                        foreach (ExtendedProperty extendedProperty in myItem.ExtendedProperties)
                                        {
                                            if (extendedProperty.PropertyDefinition.Name == filingStatus.Name &&
                                                    extendedProperty.PropertyDefinition.PropertySetId == filingStatus.PropertySetId)
                                            {
                                                myItem.RemoveExtendedProperty(filingStatus);
                                                break;
                                            }

                                            extendedPropertyindex++;
                                        }

                                        foreach (ExtendedProperty extendedProperty in myItem.ExtendedProperties)
                                        {
                                            if (extendedProperty.PropertyDefinition.Name == emailGuidProp.Name &&
                                                    extendedProperty.PropertyDefinition.PropertySetId == emailGuidProp.PropertySetId)
                                            {
                                                myItem.RemoveExtendedProperty(emailGuidProp);
                                                break;
                                            }

                                            extendedPropertyindex++;
                                        }


                                        myItem.ItemClass = "IPM.Note";
                                        //bUpdate = true;

                                        if (sReportMode.ToUpper() == "FALSE")
                                            myItem.Update(ConflictResolutionMode.AlwaysOverwrite);

                                        Log.WriteLine("Updated : Subject - {0}", (myItem as EmailMessage).Subject);
                                        Log.WriteLine("Updated : EWSId - {0}", myItem.Id.UniqueId);
                                        Log.WriteLine("");
                                        Console.WriteLine("Updated : {0}", (myItem as EmailMessage).Subject);
                                        iTotalEmailsReset++;
                                    }
                                }
                            }

                            //if (bUpdate)
                            //{                                
                            //    if (sReportMode.ToUpper() == "FALSE")
                            //        service.UpdateItems(findResults, id, ConflictResolutionMode.AlwaysOverwrite, MessageDisposition.SaveOnly, null);                                

                            //}
                            view.Offset += 50;
                        } while (findResults.MoreAvailable);

                    }
                    bRet = true;
                } while (false);
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("The specified object was not found in the store"))
                    Log.WriteLine(ex.Message);
                else
                    Log.WriteLine(ex.Message + ex.StackTrace);

                bRet = false;
            }
            Log.WriteLine("\n------------------------------------------------------------------------------------------------\n");
            Log.WriteLine("");
            Log.WriteLine("Total emails reset for {0} on Folder: {1} - {2} - {3}", sUserSmtp, sOutlookFolderName, sFolderEwsId, iTotalEmailsReset);
            Log.WriteLine("");
            Log.WriteLine("Total emails skipped for {0} on Folder: {1} - {2} - {3}", sUserSmtp, sOutlookFolderName, sFolderEwsId, iSkippedQueuedEMails);
            Log.WriteLine("");

            return bRet;
        }

        public void ScanLinkedFoldersForFilingStatusFiledMsgClsQueued(string[] args)
        {

            if (args.Length < 9)
            {
                Console.WriteLine("Syntax: <Command> <ImpersonatorSMTP> <password> <end user> <exchange server name> <csvFile> <Start Date> <End Date> <ReportMode>");
                // Console.WriteLine("Example: SCAN-LINKED-FOLDERS ImpersonatorSMTPAddress@dev.local password user1@dev.local xchange.dev.local LinkedFolderList.csv IPM.Note True");
                return;
            }

            int iTotalCnt = 0;
            long iSkippedQueuedEMails = 0;
            long iTotalEmailsReset = 0;
            do
            {
                Log.AutoFlush = true;

                m_oEntryId = new Dictionary<String, String>();

                if (!LoadEMProjectsEx(args[3], "WO-33282-2.csv", 2)) // 2- extract rows which has valid EntryId
                {
                    Console.WriteLine(String.Format("Failed to load file WO-33282-2.csv"));
                    Log.WriteLine(String.Format("Failed to load file WO-33282-2.csv"));
                    break;
                }

                Dictionary<String, String> oFolderEntryIds = null;
                oFolderEntryIds = new Dictionary<String, String>();

                //StreamWriter Log = new StreamWriter("ScanLinkedFoldersOutput.txt", true);
                //Log.AutoFlush = true;

                // Create the binding.
                ExchangeService service;
                service = new ExchangeService(ExchangeVersion.Exchange2010);


                service.Credentials = new WebCredentials(args[1], args[2]);
                service.TraceListener = new TraceListener();
                service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

                string smtpAddress = args[3];
                service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

                String sExchServer = args[4];
                String sExchSrv;
                string[] exchArr = sExchServer.Split('>');
                if (exchArr.Count() > 1)
                    sExchSrv = exchArr[1];
                else if (sExchServer.Length > 0)
                    sExchSrv = sExchServer;
                else
                {
                    Log.WriteLine("Exchange server field is blank");
                    break;
                }

                string exchangeUrl;
                exchangeUrl = "https://";
                exchangeUrl += sExchSrv;//args[4];
                exchangeUrl += "/EWS/Exchange.asmx";


                service.Url = new Uri(exchangeUrl);


                Console.WriteLine("AutodiscoverURL: " + service.Url);

                ServicePointManager.ServerCertificateValidationCallback = Program.CertificateValidationCallback;

                service.TraceEnabled = true;

                String CSVFilePath = args[5];
                if (!File.Exists(CSVFilePath))
                {
                    Log.WriteLine("File doesn't exist - {0}", CSVFilePath);
                    Console.WriteLine("File doesn't exist - {0}", CSVFilePath);
                    break;
                }

                System.IO.StreamReader file = new System.IO.StreamReader(CSVFilePath);

                string line;
                while ((line = file.ReadLine()) != null)
                {
                    line.Trim();
                    String[] Tokens = line.Split(",".ToCharArray());
                    if (2 > Tokens.Length)
                    {
                        Log.WriteLine("Invalid entry in {0}", CSVFilePath);
                        break;
                    }
                    if (!oFolderEntryIds.ContainsKey(Tokens[1].ToString()))
                        oFolderEntryIds.Add(Tokens[1].ToString(), Tokens[0].ToString());
                    else
                        Console.WriteLine("Record Exist");
                }

                foreach (KeyValuePair<String, String> Entry in oFolderEntryIds)
                {
                    try
                    {
                        String FoldEwsId = Entry.Key;
                        String FoldName = Entry.Value;
                        int iCount = 0;


                        Folder fld;
                        FolderId id = new FolderId(FoldEwsId);

                        fld = Folder.Bind(service, id);
                        Console.WriteLine("Folder Name: " + fld.DisplayName);
                        FoldName = fld.DisplayName;

                        SearchFilter.SearchFilterCollection searchAndFilterCollection =
                                                new SearchFilter.SearchFilterCollection(LogicalOperator.And);

                        //SearchFilter.SearchFilterCollection searchOrFilterCollection =
                        //                        new SearchFilter.SearchFilterCollection(LogicalOperator.Or);

                        //ExtendedPropertyDefinition emailGuidProp = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders,
                        //                                                                       "x-autn-guid", MapiPropertyType.String);

                        ExtendedPropertyDefinition filingStatus = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                                    "FilingStatus", MapiPropertyType.String);


                        ExtendedPropertyDefinition lastChangeDt = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                    "x-autn-lastchange-time", MapiPropertyType.SystemTime);


                        String sStartDt = args[6];
                        String sEndDt = args[7];
                        //if (sStartDt == "")
                        //    sStartDt = "2014-10-01";
                        //if (sEndDt == "")
                        //    sEndDt = "2015-07-01";

                        String sReportMode = args[8];

                        //searchFilterCollection.Add(new SearchFilter.IsGreaterThan(lastChangeDt, sStartDt));
                        //searchFilterCollection.Add(new SearchFilter.IsGreaterThan(EmailMessageSchema.DateTimeReceived, sStartDt));

                        //searchFilterCollection.Add(new SearchFilter.IsLessThan(lastChangeDt, sEndDt));
                        //searchFilterCollection.Add(new SearchFilter.IsLessThan(EmailMessageSchema.DateTimeReceived, sEndDt));

                        //if (!File.Exists("MessageClasses.txt"))
                        //{
                        //    Log.WriteLine("File doesn't exist - MessageClasses.txt");
                        //    Console.WriteLine("File doesn't exist - MessageClasses.txt");
                        //    break;
                        //}

                        //System.IO.StreamReader fileMsgClass = new System.IO.StreamReader("MessageClasses.txt");

                        //string lineMsgCls;
                        //while ((lineMsgCls = fileMsgClass.ReadLine()) != null)
                        //{
                        //    lineMsgCls.Trim();
                        //    if (lineMsgCls.Length > 1)
                        //        searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, lineMsgCls));
                        //}

                        //searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.WorkSite.Ems.Queued"));//args[6]));//"IPM.Note.WorkSite.Ems.Filed"));
                        //searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.ABCD"));
                        searchAndFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.WorkSite.Ems.Queued"));
                        searchAndFilterCollection.Add(new SearchFilter.IsEqualTo(filingStatus, "Filed"));
                        //searchAndFilterCollection.Add(new SearchFilter.Exists(emailGuidProp));

                        if (sStartDt.Length > 1)
                        {
                            sStartDt += "T00:00:00Z";
                            searchAndFilterCollection.Add(new SearchFilter.IsGreaterThan(lastChangeDt, sStartDt));
                        }
                        if (sEndDt.Length > 1)
                        {
                            sEndDt += "T23:59:50Z";
                            searchAndFilterCollection.Add(new SearchFilter.IsLessThan(lastChangeDt, sEndDt));
                        }


                       // searchAndFilterCollection.Add(searchOrFilterCollection);

                        FindItemsResults<Item> findResults;
                        ItemView view = new ItemView(100, 0, OffsetBasePoint.Beginning);

                        // Identify the Subject properties to return.
                        // Indicate that the base property will be the item identifier
                        view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.ItemClass, filingStatus);//emailGuidProp, filingStatus);

                        // Set the traversal to shallow. (Shallow is the default option; other options are Associated and SoftDeleted.)
                        view.Traversal = ItemTraversal.Shallow;



                        do
                        {


                            // Send the request to search the Inbox and get the results.
                            findResults = service.FindItems(id, searchAndFilterCollection, view);
                            //findResults = service.FindItems(id, view);

                            bool bUpdate = true;

                            if (bUpdate)
                            {
                                //int extendedPropertyindex = 0;
                                //bool bUpdate = false;

                                // Process each item.
                                foreach (Item myItem in findResults.Items)
                                {
                                    //extendedPropertyindex = 0;

                                    if (myItem is EmailMessage)
                                    {
                                        // Get EntryId from EWSId

                                        AlternateId objAltID = new AlternateId();
                                        objAltID.Format = IdFormat.EwsId;
                                        objAltID.Mailbox = smtpAddress;
                                        objAltID.UniqueId = myItem.Id.ToString();

                                        //Convert  PR_ENTRYID identifier format to an EWS identifier. 
                                        AlternateIdBase objAltIDBase = service.ConvertId(objAltID, IdFormat.HexEntryId);
                                        AlternateId objAltIDResp = (AlternateId)objAltIDBase;

                                        // Check if this EntryId present in the EM_REQUEST
                                        if (m_oEntryId.ContainsKey(objAltIDResp.UniqueId))
                                        {
                                            Log.WriteLine("Skip - EntryId {0} exist in EM_REQUEST", objAltIDResp.UniqueId);
                                            iSkippedQueuedEMails++;
                                            continue;
                                        }
                                        else
                                        {
                                            //foreach (ExtendedProperty extendedProperty in myItem.ExtendedProperties)
                                            //{
                                            //    if (extendedProperty.PropertyDefinition.Name == filingStatus.Name &&
                                            //            extendedProperty.PropertyDefinition.PropertySetId == filingStatus.PropertySetId)
                                            //    {
                                            //        myItem.RemoveExtendedProperty(filingStatus);
                                            //        break;
                                            //    }

                                            //    extendedPropertyindex++;
                                            //}

                                            //foreach (ExtendedProperty extendedProperty in myItem.ExtendedProperties)
                                            //{
                                            //    if (extendedProperty.PropertyDefinition.Name == emailGuidProp.Name &&
                                            //            extendedProperty.PropertyDefinition.PropertySetId == emailGuidProp.PropertySetId)
                                            //    {
                                            //        myItem.RemoveExtendedProperty(emailGuidProp);
                                            //        break;
                                            //    }

                                            //    extendedPropertyindex++;
                                            //}


                                            //if ((myItem.ItemClass.ToUpper() == "IPM.NOTE.WORKSITE.EMS.QUEUED") ||
                                            //    (myItem.ItemClass.ToUpper() == "IPM.NOTE.WORKSITE.EMS.FILED"))
                                            //    myItem.ItemClass = "IPM.Note";

                                           // bUpdate = true;

                                            //if (sReportMode.ToUpper() == "FALSE")
                                            //    myItem.Update(ConflictResolutionMode.AlwaysOverwrite);

                                            Log.WriteLine("Updated : Subject - {0}", (myItem as EmailMessage).Subject);
                                            Log.WriteLine("Updated : EWSId - {0}", myItem.Id.UniqueId);
                                            Log.WriteLine("");
                                            Console.WriteLine("Updated : {0}", (myItem as EmailMessage).Subject);
                                            iTotalEmailsReset++;
                                        }
                                    }
                                }
                            }

                            // Process each item.
                            //foreach (Item myItem in findResults.Items)
                            //{
                            //    if (myItem is EmailMessage)
                            //    {
                            //        iCount++;
                            //        Console.WriteLine((myItem as EmailMessage).Subject);
                            //    }
                            //}
                            iCount += findResults.Items.Count();
                            iTotalCnt += findResults.Items.Count();
                            view.Offset += 100;
                        } while (findResults.MoreAvailable);


                        Console.WriteLine("Reset count : {0} ", iTotalEmailsReset);
                        //Log.WriteLine("Folder : {0} : Items Processed : {1}", FoldName, iTotalCnt);

                        Log.WriteLine("");
                        Log.WriteLine("Total emails reset for {0} on Folder: {1} - {2} - {3}", smtpAddress, FoldName, FoldEwsId, iTotalEmailsReset);
                        Log.WriteLine("");
                        Log.WriteLine("Total emails skipped (request exist in em_req) for {0} on Folder: {1} - {2} - {3}", smtpAddress, FoldName, FoldEwsId, iSkippedQueuedEMails);
                        Log.WriteLine("");

                    }
                    catch (Exception ex)
                    {
                        Log.WriteLine("Folder: {0} : ", ex.Message);
                    }
                }
                Log.WriteLine("Total Items Processed : {0} ", iTotalCnt);
            } while (false);


        }
        
    }
}
