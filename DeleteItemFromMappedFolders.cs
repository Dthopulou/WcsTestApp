using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using Com.Interwoven.WorkSite.iManage;
using System.Threading.Tasks;

namespace EWSTestApp
{
    class DeleteItemFromMappedFolders
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
        public DeleteItemFromMappedFolders(ref WorkSiteUtility workUtility)
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

 public void DeleteFiledItemsFromMappedFolder(string[] args)
        {
            if (args.Length < 9)
            {
                Console.WriteLine("Syntax: <Command> <WorkServer> <NRTAdmin> <password> <ImpersonatorSMTP> <password> <exchange server name> <ExchangeVersion> <sReportmode>");
                Console.WriteLine("Example: GET_ALL_FOLDER_MAPPINGS WorkSite NRTAdmin password ImpersonatorSMTPAddress@dev.local password exchangeServer Exchange2010 False/True");
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
                    break;
                }

                string sImpersonationAC = args[4];
                string sPassword = args[5];
                string sExchServer = args[6];
                string sExchVersion = args[7];
                string sReportMode = args[8].ToUpper();
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
                            //Get All the mapped folders

                            if (m_oWorkSession.GetMappedFolders(sUser))
                            {
                                Dictionary<String, FolderMapping> oEmFolderMappings = null;

                                m_oWorkSession.GetMappedFolderCollection(ref oEmFolderMappings);

                                if (oEmFolderMappings == null)
                                {
                                    Log.WriteLine("No mapped folder for user {0}", smtpAddress);
                                }

                                else
                                {
                                    foreach (KeyValuePair<String, FolderMapping> folderMap in oEmFolderMappings)
                                    {
                                        if (null == oEmFolderMappings)
                                        { continue; }

                                        FolderMapping value = folderMap.Value;
                                        // calling IsDeleteMessageSet 
                                        if (!m_oWorkSession.IsDeleteMessageSet(value.OtherProperties))
                                        {
                                            continue;
                                        }

                                        string sMsg = "";
                                        string foldName = "";
                                        try
                                        {
                                            foldName = "";
                                            sMsg = "";
                                            String FoldEwsId;
                                            FoldEwsId = ConvertID(ref service, smtpAddress, "HEX", "EWSID", folderMap.Value.exchFolderID);
                                            if ((FoldEwsId != null) && (FoldEwsId.Length > 0))
                                            {
                                                Folder fld;
                                                FolderId id = new FolderId(FoldEwsId);
                                                fld = Folder.Bind(service, id);
                                                FindItemsResults<Item> allItemsInThisFolder;

                                                int iBatchSize = 100;
                                                ItemView view1 = new ItemView(iBatchSize, 0, OffsetBasePoint.Beginning);
                                                view1.Traversal = ItemTraversal.Shallow;
                                                view1.PropertySet = new PropertySet(BasePropertySet.FirstClassProperties);

                                                ExtendedPropertyDefinition filingStatus = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                                                         "FilingStatus", MapiPropertyType.String);

                                                SearchFilter searchFilter = new SearchFilter.IsEqualTo(filingStatus, "Filed");
                                                if (sReportMode == "TRUE")
                                                {

                                                    Log.WriteLine("this utility will delete all the filed emails in the mapped folder when the users uncheck the 'leave message in outlook'");
                                                }
                                                else
                                                {
                                                    do
                                                    {
                                                        //find all the filed items within the mapped folder whose filingstatus is filed and delete them

                                                        allItemsInThisFolder = service.FindItems(id, searchFilter, view1);

                                                        if (allItemsInThisFolder.Count() == 0)
                                                        {
                                                            Log.WriteLine("The folder {0} does not have any items", fld.DisplayName);
                                                        }
                                                        else
                                                        {
                                                            foreach (Item item in allItemsInThisFolder.Items)
                                                            {
                                                                item.Delete(DeleteMode.HardDelete);
                                                            }
                                                        }


                                                    } while (allItemsInThisFolder.MoreAvailable);

                                                    Log.WriteLine("Total number of items deleted from folder {0} is: {1}", fld.DisplayName, allItemsInThisFolder.Count());
                                                }


                                                if (fld != null)
                                                {
                                                    foldName = fld.DisplayName;
                                                    foldName = foldName.Replace(',', '_');
                                                }
                                            }

                                            sMsg += folderMap.Value.databaseName;
                                            sMsg += ",";
                                            sMsg += folderMap.Value.sid;
                                            sMsg += ",";
                                            sMsg += folderMap.Value.prjID.ToString();
                                            sMsg += ",";
                                            sMsg += folderMap.Value.userID;
                                            sMsg += ",";
                                            sMsg += folderMap.Value.status.ToString();
                                            sMsg += ",";

                                            string sDesc = folderMap.Value.statusDescription;
                                            if (sDesc != null)
                                                sDesc = sDesc.Replace(',', '_');
                                            else
                                                sDesc = "";

                                            sMsg += sDesc;
                                            sMsg += ",";
                                            sMsg += folderMap.Value.isActive.ToString();
                                            sMsg += ",";
                                            sMsg += folderMap.Value.sOperator;
                                            sMsg += ",";
                                            sMsg += folderMap.Value.lastSyncTime;
                                            sMsg += ",";

                                            string sFoldPath = folderMap.Value.projectFolderPath;
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
                                            sMsg += folderMap.Value.foldEntryId;
                                            EwsMappedFolderReportLog.WriteLine(sMsg);
                                        }
                                        catch (Exception ex)
                                        {
                                            DateTime dt = DateTime.Now;
                                            //String.Format("{0:u}", dt);
                                            Log.WriteLine("ScanAllFolderMapping report: {0}, Message: {1} EntryId: {2}, SID: {3} ", dt, ex.Message, folderMap.Key, folderMap.Value.sid);
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

    }
}
