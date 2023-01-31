using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Com.Interwoven.WorkSite.iManage;
using System.Data.SqlClient;
using System.IO;

namespace EWSTestApp
{
    class WorkSiteUtility
    {
        private NRTDMS nrtdms;

        private Dictionary<String, ExplicitRequest> m_oEmRequests = null;
        private Dictionary<String, FolderMapping> m_oEmFolderMappings = null;
        
        
        

        public WorkSiteUtility()
        {
            nrtdms = new NRTDMS();
            
        }

       
        public bool Login(string serverName_, string userID_, string password_)
        {
            bool bRet = false;
            if (nrtdms.Sessions.Count > 0)
            {
                for (int i = 1; i <= nrtdms.Sessions.Count; i++)
                {
                    nrtdms.Sessions.Remove(i);
                }
            }

            INRTSession session = nrtdms.Sessions.Add(serverName_);

            try
            {
                session.Login(userID_, password_);
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);                
            }

            // Report a successful login attempt if there was one.
            if (session.Connected)
            {
                bRet = true;
            }
            return bRet;
        }

        public bool IsFolderMappingExist(string parentFolderEntryId, out FolderMapping exReq)
        {
            bool bRet = false;
            exReq = null;
            if (m_oEmFolderMappings.ContainsKey(parentFolderEntryId))
            {
                if (m_oEmFolderMappings.TryGetValue(parentFolderEntryId, out exReq))
                    bRet = true;
            }

            return bRet;
        }

        //checking message-delete other properties
        public bool IsDeleteMessageSet(ManStrings folderMapProperties)
        {
            bool isDeleteMessageSet = true;
            for (int nIndex = 1; nIndex <= folderMapProperties.Count; nIndex++)
            {
                string folderMapProperty = folderMapProperties.ItemByIndex(nIndex);
                folderMapProperty = folderMapProperty.ToLower();
                if (folderMapProperty.StartsWith("message-delete"))
                {
                    int nValueIndex = folderMapProperty.IndexOf('=') + 1;
                    string deleteMessage = folderMapProperty.Substring(nValueIndex);
                    isDeleteMessageSet = (deleteMessage == "y") || (deleteMessage == "1");
                    break;
                }
            }
            return isDeleteMessageSet;
        }

        public bool IsEmailRequestExists(ExchangeQueuedEmails exchangeEmail, out ExplicitRequest exReq)
        {
            bool bRet = false;
            exReq = null;
            if (m_oEmRequests.ContainsKey(exchangeEmail.entryId))
            {
                if (m_oEmRequests.TryGetValue(exchangeEmail.entryId, out exReq))
                    bRet = true;
            }

            return bRet;
        }

        public void GetMappedFolderCollection(ref Dictionary<String, FolderMapping> oEmFolderMappings)
        {
            if (m_oEmFolderMappings.Count() > 0)
                oEmFolderMappings = m_oEmFolderMappings;
        }

        public bool GetMappedFolders(string userID_)
        {
            bool bRet = false;
            try
            {
                // Retrieve the explicit filing requests across databases on the server.
                IManDMS imanDMS = (IManDMS)nrtdms;

                for (int iDbCnt = 1; iDbCnt <= imanDMS.Sessions.Count; iDbCnt++)
                {
                    IManSession session = imanDMS.Sessions.ItemByIndex(iDbCnt);

                    IManDatabase database = null;
                    IManDatabaseEM2 databaseEM2 = null;

                    // Get the filing requests from the database, filtering by user if specified.
                    try
                    {
                        database = session.Databases.ItemByIndex(iDbCnt);
                    }
                    catch (Exception except)
                    {
                        Console.WriteLine("Error: {0}, {1}", session.ServerName, except.Message);
                    }
                    databaseEM2 = (Com.Interwoven.WorkSite.iManage.IManDatabaseEM2)database;
                    IEMFolderMappings emFolderMappings = null;
                    try
                    {
                        emFolderMappings = databaseEM2.EMFolderMappingsForUser(userID_);//EMRequestStatus.EMRequestFailure);                        
                    }
                    catch (COMException except)
                    {
                        Console.WriteLine("Error: {0}", except.Message);
                    }
                    catch (Exception except)
                    {
                        Console.WriteLine("Error: {0}", except.Message);
                    }

                    // Package the filing request information relevant to this application.
                    if (emFolderMappings != null)
                    {
                        try
                        {
                            if (m_oEmFolderMappings == null)
                                m_oEmFolderMappings = new Dictionary<String, FolderMapping>();
                            foreach (IEMFolderMapping emFolderMapping in emFolderMappings)
                            {
                                if (!m_oEmFolderMappings.ContainsKey(emFolderMapping.EMFolder))
                                {
                                    FolderMapping foldMapping = new FolderMapping(emFolderMapping);
                                    m_oEmFolderMappings.Add(emFolderMapping.EMFolder, foldMapping);
                                }
                            }

                            bRet = true;


                        }
                        catch (Exception except)
                        {
                            Console.WriteLine("Error: {0}", except.Message);
                        }
                    }
                }
                
            }            
            catch (Exception except)
            {
                Console.WriteLine("Error: {0}", except.Message);   
            }

            return bRet;
        }


        public bool InsertEMRequestEntry(string sDatabaseName, ExchangeQueuedEmails queuedEmailExch, ref StreamWriter Log)
        {
            bool bRet = false;
            try
            {

                if ((sDatabaseName == null) || (sDatabaseName.Length == 0))
                {
                    Log.WriteLine("InsertEMRequest Invalid parameter :{0} ", sDatabaseName);
                    return bRet;
                }

                // Retrieve the explicit filing requests across databases on the server.
                IManDMS imanDMS = (IManDMS)nrtdms;

                for (int iDbCnt = 1; iDbCnt <= imanDMS.Sessions.Count; iDbCnt++)
                {
                    IManSession session = imanDMS.Sessions.ItemByIndex(iDbCnt);

                    IManDatabase database = null;
                    IManDatabaseEM2 databaseEM2 = null;

                    // Get the filing requests from the database, filtering by user if specified.
                    try
                    {
                        database = session.Databases.ItemByIndex(iDbCnt);
                        if (database.Name != sDatabaseName)
                        {
                            Log.WriteLine("InsertEMRequestEntry : Database: {0}", database.Name);
                            continue;
                        }

                    }
                    catch (Exception except)
                    {
                        Log.WriteLine("InsertEMRequest: {0}, {1}", session.ServerName, except.Message);
                    }
                    databaseEM2 = (Com.Interwoven.WorkSite.iManage.IManDatabaseEM2)database;
                    IEMFilingRequest emFilingRequest = null;
                    string sWorkUser = "";
                    string sOtherProperty = "";
                    try
                    {
                       if (GetWorkSiteUser(sDatabaseName, queuedEmailExch.EmailId, ref sWorkUser))
                       {
                           emFilingRequest = databaseEM2.CreateFilingRequest();

                           emFilingRequest.UserID = sWorkUser;

                           if (queuedEmailExch.PrjId == 0)
                           {
                               Log.WriteLine("InsertEMRequest: Couldn't Work Folder ID {0}, {1}", queuedEmailExch.ewsId, queuedEmailExch.entryId);
                               break;
                           }

                           emFilingRequest.FolderID = queuedEmailExch.PrjId;
                           emFilingRequest.Mailbox = "WCSUtility";
                           emFilingRequest.EMObjectID = queuedEmailExch.entryId;
                           emFilingRequest.EMFolder = queuedEmailExch.parentFolderEntryId;
                           emFilingRequest.RetryCount = 0;
                           emFilingRequest.RequestType = EMFilingRequestType.FilingRequestFile;
                           emFilingRequest.ClientType = EMClientType.ClientMSOutlook;
                           emFilingRequest.StatusCode = EMRequestStatus.EMRequestSubmitted;
                           emFilingRequest.Operation = EMOperation.OperationCopy;

                           sOtherProperty = "PR_OPERATOR=" +sWorkUser;
                           emFilingRequest.OtherProperties.Add(sOtherProperty);

                           sOtherProperty = "email-searchkey=" + queuedEmailExch.searchKey;
                           emFilingRequest.OtherProperties.Add(sOtherProperty);

                           emFilingRequest.Update();
                           bRet = true;
                       }
                       else
                           Log.WriteLine("InsertEMRequest: Couldn't get Work user for {0}", queuedEmailExch.EmailId);
                    }
                    catch (COMException except)
                    {
                        Log.WriteLine("InsertEMRequestEntry: {0}", except.Message);
                    }
                    catch (Exception except)
                    {
                        Log.WriteLine("InsertEMRequestEntry: {0}", except.Message);
                    }

                    
                }

            }
            catch (Exception except)
            {
                Console.WriteLine("InsertEMRequestEntry: {0}", except.Message);
            }

            return bRet;
        }

        public bool GetWorkSiteUser(string sDatabase, string sEmailAddress, ref string sWorkUser)
        {
            bool bRet = false;
            try
            {
                // Retrieve the explicit filing requests across databases on the server.
                IManDMS imanDMS = (IManDMS)nrtdms;

                for (int iDbCnt = 1; iDbCnt <= imanDMS.Sessions.Count; iDbCnt++)
                {
                    IManSession session = imanDMS.Sessions.ItemByIndex(iDbCnt);

                    IManDatabase database = null;
                    
                    try
                    {
                        database = session.Databases.ItemByIndex(iDbCnt);

                        if (database.Name != sDatabase)
                            continue;
                    }
                    catch (Exception except)
                    {
                        Console.WriteLine("Error: {0}, {1}", session.ServerName, except.Message);
                    }

                    NRTUserParameters userParameters = (NRTUserParameters)nrtdms.CreateUserParameters();

                    if (userParameters != null)
                    {
                        userParameters.Add(UserAttributeID.nrUserEmail, sEmailAddress);

                        INRTDatabase nrtDb = (INRTDatabase)database;

                        if (nrtDb != null)
                        {
                            Object oFlow = new Object();
                            NRTUsers users = (NRTUsers)nrtDb.FindUsers(userParameters, oFlow);
                            if (users != null)
                            {
                                NRTUser user = users.Item(1);
                                sWorkUser = user.Name;
                                Console.WriteLine(user.Name);
                                bRet = true;
                                break;
                            }
                        }
                    }
                }

            }
            catch (Exception except)
            {
                Console.WriteLine("Error: {0}", except.Message);
            }

            return bRet;
        }
        public bool GetExplicitRequests(/*string libraryName_,*/ string userID_, EMRequestStatus status)
        {
            bool bRet = false;
            try
            {
                // Retrieve the explicit filing requests across databases on the server.
                IManDMS imanDMS = (IManDMS)nrtdms;

                for (int iDbCnt = 1; iDbCnt <= imanDMS.Sessions.Count; iDbCnt++)
                {
                    IManSession session = imanDMS.Sessions.ItemByIndex(iDbCnt);

                    IManDatabase database = null;
                    IManDatabaseEM2 databaseEM2 = null;

                    // Get the filing requests from the database, filtering by user if specified.
                    try
                    {
                        database = session.Databases.ItemByIndex(iDbCnt);
                    }
                    catch (Exception except)
                    {
                        Console.WriteLine("Error: {0}, {1}", session.ServerName, except.Message);
                    }
                    databaseEM2 = (Com.Interwoven.WorkSite.iManage.IManDatabaseEM2)database;
                    IEMFilingRequests emFilingRequests = null;
                    try
                    {
                        if (String.IsNullOrEmpty(userID_))
                        {
                            emFilingRequests = databaseEM2.EMFilingRequestsForAllUsersByStatusCode(status);//EMRequestStatus.EMRequestFailure);

                        }
                        else
                        {
                            emFilingRequests = databaseEM2.EMFilingRequestsForUserByStatusCode(userID_, status);//EMRequestStatus.EMRequestFailure);
                        }
                    }
                    catch (COMException except)
                    {
                        Console.WriteLine("Error: {0}", except.Message);
                    }
                    catch (Exception except)
                    {
                        Console.WriteLine("Error: {0}", except.Message);
                    }

                    // Package the filing request information relevant to this application.
                    if (emFilingRequests != null)
                    {
                        try
                        {
                            if (m_oEmRequests == null)
                                m_oEmRequests = new Dictionary<String, ExplicitRequest>();
                            foreach (IEMFilingRequest emFilingRequest in emFilingRequests)
                            {
                                if (!m_oEmRequests.ContainsKey(emFilingRequest.EMObjectID))
                                {
                                    ExplicitRequest explicitRequest = new ExplicitRequest(emFilingRequest);
                                    m_oEmRequests.Add(emFilingRequest.EMObjectID, explicitRequest);
                                }
                            }

                            bRet = true;


                        }
                        catch (Exception except)
                        {
                            Console.WriteLine("Error: {0}", except.Message);
                        }
                    }
                }
                
            }            
            catch (Exception except)
            {
                Console.WriteLine("Error: {0}", except.Message);   
            }

            return bRet;
        }

        public void ClearEmAndMappingRequests()
        {
            if (m_oEmRequests != null)
                m_oEmRequests.Clear();
            m_oEmRequests = null;

            if (m_oEmFolderMappings != null)
                m_oEmFolderMappings.Clear();
            m_oEmFolderMappings = null;
        }

        public bool GetWhereFiledInfo(double lDocNum, int lDocVer, ref string sPath, ref string sServerName, ref string sDb, ref int iPrjId, ref System.IO.StreamWriter Log)
        {
            //CheckEmailExistInWorkServerDatabase(ref msgIdFromExchange, ref msgIdInWorkServer, ref Log);
            bool bRet = false;
            string workingPath = String.Empty;
            try
            {
                IManDMS imanDMS = (IManDMS)nrtdms;
                IManProfileSearchParameters imanProfileParameters = imanDMS.CreateProfileSearchParameters();
                IManSession session = imanDMS.Sessions.ItemByIndex(1);
                ManStrings strDbs = new Com.Interwoven.WorkSite.iManage.ManStrings();
                iPrjId = 0;
                for (int iDbCnt = 1; iDbCnt <= session.Databases.Count; iDbCnt++)
                {
                    workingPath = "";
                    IManDatabase database = null;

                    database = session.Databases.ItemByIndex(iDbCnt);

                    IManDocument doc = database.GetDocument((int)lDocNum, lDocVer);
                    if (doc != null)
                    {
                        IManFolders folds = doc.Folders;

                        IManFolder parent = null;
                        // The use of the last folder variable is intended to build the path of folders without the name (last element).

                        foreach (IManFolder currentFolder in folds)
                        {
                            if (iPrjId == 0)
                                iPrjId = currentFolder.FolderID;
                            parent = currentFolder;
                            workingPath += "\\" + currentFolder.Name;
                        }
                        while ((parent = parent.Parent) != null)
                        {
                            workingPath = "\\" + parent.Name + workingPath;
                        }
                        sPath = "\\" + workingPath;
                        sServerName = session.ServerName;
                        sDb = doc.Database.Name;
                        bRet = true;
                        break;
                    }

                }

            }
            catch (Exception except)
            {
                Log.WriteLine("Error GetWhereFiledInfo : {0}", except.Message);
                bRet = false;
            }

            return bRet;
        }

        public int CheckEmailExistInWorkServerDatabase(ref Dictionary<string, string> oDbConns, ref List<string> msgIdFromExchange, ref List<FiledEmailDetails> msgIdInWorkServer, ref System.IO.StreamWriter Log)
        {
            int iRet = 0;

            foreach (KeyValuePair<String, string> conn in oDbConns)
            {
                iRet = 0;
                try
                {
                    string connetionString = null;
                    SqlConnection cnn;
                    SqlCommand command;

                    //connetionString = "Data Source=10.192.211.228;Initial Catalog=WS_DB_94_1;User ID=SA;Password=Password1";
                    connetionString = conn.Value;

                    if (connetionString.Length <= 0)
                        continue;

                    cnn = new SqlConnection(connetionString);
                   // string sql = null;
                    SqlDataReader dataReader;
                    //sql = "select * from MHGROUP.DOCMASTER (nolock) where MSG_ID in ('6f84c287f67a4c4d90644aa24d51aad3@exdev2016.local', 'ecab690601534e3a84d44715b71fbe0e@exdev2016.local'";
                    //sql = "select MSG_ID from MHGROUP.DOCMASTER (nolock) where MSG_ID in ('6f84c287f67a4c4d90644aa24d51aad3@exdev2016.local','ecab690601534e3a84d44715b71fbe0e@exdev2016.local')";//;

                    string strMsgIds = "";
                    foreach (string sId in msgIdFromExchange)
                    {
                        strMsgIds += "'";
                        strMsgIds += sId;
                        strMsgIds += "'";
                        strMsgIds += ",";
                    }
                    if (strMsgIds.Length > 0)
                        strMsgIds = strMsgIds.Substring(0, strMsgIds.Length - 1);

                    string sSearchCriteria = "select MSG_ID, DOCNUM, VERSION from MHGROUP.DOCMASTER (nolock) where MSG_ID in (";
                   // sSearchCriteria += "'ecab690601534e3a84d44715b71fbe0e@exdev2016.local', '484fdb2842b6414ab42f71858877d664@exdev2016.local', 'd0a7cf548fb84e7f9a1d347993ef1da7@exdev2016.local'";//strMsgIds;
                    sSearchCriteria += strMsgIds;
                    sSearchCriteria += ")";
                    //"(IM_MSGID:(b4c4bb42-cca0-4ce6-8691-e0661e6ca888@HobbitExch.shire.local))"
                    //imanProfileParameters.AddFullTextSearch(sSearchCriteria, imFullTextSearchLocation.imFullTextAnywhere);
                    //imanProfileParameters.Add(imFullTextSearchLocation.imFullTextAnywhere, "(IM_MSGID:(b4c4bb42-cca0-4ce6-8691-e0661e6ca888@HobbitExch.shire.local))");
                    //(IM_DOCNAME:(test))


                    cnn.Open();
                    Log.WriteLine("Connection Open ! ");
                    command = new SqlCommand(sSearchCriteria, cnn);
                    dataReader = command.ExecuteReader();
                    //double db = 0;
                    double f1 = 0;
                    int i1 = 0;
                    while (dataReader.Read())
                    {
                       // Log.WriteLine(dataReader.GetValue(0));// + " - " + dataReader.GetValue(1) + " - " + dataReader.GetValue(2));
                        FiledEmailDetails oFiledEmail = new FiledEmailDetails();
                        oFiledEmail.messageId = (string)dataReader.GetValue(0);
                        oFiledEmail.DocNum = dataReader.GetDouble(1);
                        oFiledEmail.Version = dataReader.GetInt32(2);
                        msgIdInWorkServer.Add(oFiledEmail);
                        //msgIdInWorkServer.Add((string)dataReader.GetValue(0));
                        iRet = 1;
                    }
                    dataReader.Close();
                    cnn.Dispose();

                    cnn.Close();
                }
                catch (COMException comExcep)
                {
                    Log.WriteLine("Com Error: {0}", comExcep.Message);
                    iRet = 2;
                }
                catch (Exception except)
                {
                    Log.WriteLine("Can not open connection - Check your credentials ! {0}", except.Message);
                    iRet = 2;
                }
            }

            return iRet;
        }

        //public int CheckEmailExistInWorkServerDatabase(ref Dictionary<string, string> oDbConns, ref List<string> msgIdFromExchange, ref List<string> msgIdInWorkServer, ref System.IO.StreamWriter Log)
        //{
        //    int iRet = 0;
        //   string connetionString = null;
        //    SqlConnection cnn ;
        //    SqlCommand command;

        //    connetionString = "Data Source=10.192.211.228;Initial Catalog=WS_DB_94_1;User ID=SA;Password=Password1";
        //    cnn = new SqlConnection(connetionString);
        //    string sql = null;
        //    SqlDataReader dataReader;
        //    //sql = "select * from MHGROUP.DOCMASTER (nolock) where MSG_ID in ('6f84c287f67a4c4d90644aa24d51aad3@exdev2016.local', 'ecab690601534e3a84d44715b71fbe0e@exdev2016.local'";
        //    sql = "select MSG_ID from MHGROUP.DOCMASTER (nolock) where MSG_ID in ('6f84c287f67a4c4d90644aa24d51aad3@exdev2016.local','ecab690601534e3a84d44715b71fbe0e@exdev2016.local')";//;
           
        //    try
        //    {
        //        cnn.Open();
        //        Console.WriteLine("Connection Open ! ");
        //        command = new SqlCommand(sql, cnn);
        //        dataReader = command.ExecuteReader();
        //        while (dataReader.Read())
        //        {
        //            Console.WriteLine(dataReader.GetValue(0));// + " - " + dataReader.GetValue(1) + " - " + dataReader.GetValue(2));
        //            msgIdInWorkServer.Add((string)dataReader.GetValue(0));
        //            iRet = 1;
        //        }
        //        dataReader.Close();
        //        cnn.Dispose();
                
        //        cnn.Close();
        //    }
        //    catch (COMException comExcep)
        //    {
        //        Log.WriteLine("Com Error: {0}", comExcep.Message);
        //        iRet = 2;
        //    }
        //    catch (Exception except)
        //    {
        //        Log.WriteLine("Can not open connection ! {0}", except.Message);
        //        iRet = 2;
        //    }

        //    return iRet;
        //}
        // TryEmail
        public int CheckEmailExistInWorkServer(ref List<string> msgIdFromExchange, ref List<string> msgIdInWorkServer, ref System.IO.StreamWriter Log)
        {
            //CheckEmailExistInWorkServerDatabase(ref msgIdFromExchange, ref msgIdInWorkServer, ref Log);
             int iRet = 0;
             try
             {
                 IManContents contents = null;
                 IManDMS imanDMS = (IManDMS)nrtdms;
                 IManProfileSearchParameters imanProfileParameters = imanDMS.CreateProfileSearchParameters();
                 IManSession session = imanDMS.Sessions.ItemByIndex(1);
                 ManStrings strDbs = new Com.Interwoven.WorkSite.iManage.ManStrings();

                 for (int iDbCnt = 1; iDbCnt <= session.Databases.Count; iDbCnt++)
                 {
                    
                     IManDatabase database = null;

                     database = session.Databases.ItemByIndex(iDbCnt);

                     strDbs.Add(database.Name);
                 }

                 string strMsgIds = "";
                 foreach (string sId in msgIdFromExchange)
                 {
                     strMsgIds += sId;
                     strMsgIds += ",";
                 }
                 if (strMsgIds.Length > 0)
                     strMsgIds = strMsgIds.Substring(0, strMsgIds.Length - 1);
                 
                 string sSearchCriteria = "(IM_MSGID:(";
                 sSearchCriteria += strMsgIds;
                 sSearchCriteria += "))";
                 //"(IM_MSGID:(b4c4bb42-cca0-4ce6-8691-e0661e6ca888@HobbitExch.shire.local))"
                 imanProfileParameters.AddFullTextSearch(sSearchCriteria, imFullTextSearchLocation.imFullTextAnywhere);
                 //imanProfileParameters.Add(imFullTextSearchLocation.imFullTextAnywhere, "(IM_MSGID:(b4c4bb42-cca0-4ce6-8691-e0661e6ca888@HobbitExch.shire.local))");
                 //(IM_DOCNAME:(test))


                 int iCount = 0;
                 while (iCount < 3)
                 {
                     try
                     {
                         iCount++;
                         contents = session.SearchDocuments(strDbs, imanProfileParameters, true);

                         if (contents != null)
                         {
                             foreach (IManContent content in contents)
                             {
                                 Object obj = null;
                                 IManDocument doc = (IManDocument)content;
                                 obj = doc.GetAttributeValueByID(imProfileAttributeID.imProfileMessageUniqueID);

                                 if (obj != null)
                                 {
                                     msgIdInWorkServer.Add((string)obj);
                                     iRet = 1;
                                 }

                                 
                                 //string msgID = (string) doc.GetAttributeByID(imProfileAttributeID.imProfileMessageUniqueID);

                                 //string sMsgId = (string)doc.GetAttributeByID(imProfileAttributeID.imProfileMessageUniqueID);
                                 //Console.WriteLine(doc.Description);
                                 //Console.WriteLine(doc.EditDate);
                                 //Console.WriteLine(doc.CustomAttributes);

                             }
                         }
                         iCount = 3;
                     }
                     catch (COMException comExcep)
                     {
                         Log.WriteLine("Com Error: {0}", comExcep.Message);
                         iRet = 2;
                     }
                     catch (Exception except)
                     {
                         Log.WriteLine("Error: {0}", except.Message);
                         iRet = 2;
                     }
                 }


                //string strMsgId1 = "844047319c7e4a54a464d256a75ad34d@exdev2016.local,13de38957f3e44e1a41ac1d7cf10ae50@exdev2016.local";
                //imanProfileParameters.Add(imProfileAttributeID.imProfileMessageUniqueID, strMsgIds);
                //imanProfileParameters.Add(imProfileAttributeID.imProfileDescription, "*");

                 //DateTime dt = Convert.ToDateTime("11/8/2016 9:00:47 AM");//2016-08-29 15:32:20.000");
                 //DateTime dt = Convert.ToDateTime("08/29/2016 03:32:20 PM");
                 //imanProfileParameters.Add(imProfileAttributeID.imProfileCustom21, dt.ToString());
                 //imanProfileParameters.Add(imProfileAttributeID.imProfileDescription, "send and file [IWOV-WS_DB_94_1.FID89]");
                 //imanProfileParameters.Add(imProfileAttributeID.imProfileDocNum, "207");

                 //imanProfileParameters.Add(imProfileAttributeID.imProfileDocNum, ">0");

             }
             catch (Exception except)
             {
                 Console.WriteLine("Error : {0}", except.Message);
                 iRet = 2;
             }

             return iRet;
        }

        public void test()
        {
            IManDMS imanDMS = (IManDMS)nrtdms;
            IManSession session = imanDMS.Sessions.ItemByIndex(1);
            if (session.Connected)
            {
                Console.WriteLine("Connected");
                //GetExplicitRequests("WS_DB_94_1", "", EMRequestStatus.EMRequestSubmitted);
               // GetExplicitRequests("WS_DB_94_1", "", EMRequestStatus.EMRequestFailure);

                //CheckEmailExistInWorkServer();
            }
        }
    }
}
