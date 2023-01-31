using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web.Services;
//using ProxyHelper.EWS;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.Net;
using System.Windows.Forms;
using Microsoft.Exchange.WebServices.Autodiscover;
using Microsoft.Exchange.WebServices.Data;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Xml;
using System.Collections.ObjectModel;
using Interop.MIMETranslator;

namespace EWSTestApp
{
    class TraceListener : ITraceListener
    {
        #region ITraceListener Members

        public void Trace(string traceType, string traceMessage)
        {
            CreateXMLTextFile(traceType + " - " + traceMessage.ToString());
        }

        #endregion
        private void CreateXMLTextFile(string traceContent)
        {

            string strPath = System.AppDomain.CurrentDomain.BaseDirectory;
            strPath = strPath + "\\Log.txt";
            System.IO.FileStream fs;
            if (System.IO.File.Exists(strPath) == false)
            {
                fs = System.IO.File.Create(strPath);
            }
            else
            {
                fs = System.IO.File.OpenWrite(strPath);
            }

            fs.Close();

            // Create an instance of StreamWriter to write text to a file.
            System.IO.StreamWriter sw = System.IO.File.AppendText(strPath);
            sw.WriteLine(System.DateTime.Now.ToString() + ": " + traceContent);
            sw.Close();

        }
    }

    class Program
    {
        public static bool CertificateValidationCallback(Object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }

        public static bool RedirectionUrlValidationCallback(String redirectionUrl)
        {

            return true;
        }

        

        //////Delete
        

        public static void FindLinkedFolders(string[] args)
        {

            if (args.Length < 6)
            {
                Console.WriteLine("Syntax: <Command> <ImpersonatorSMTP> <password> <end user> <exchange server name> [mapping extn]");
                Console.WriteLine("Example: GET-LINKED-FOLDERS ImpersonatorSMTPAddress@dev.local password user1@dev.local xchange.dev.local [WorkSite]");
                return;
            }

            StreamWriter Log = new StreamWriter("LinkedFolderList.csv", true);
            Log.AutoFlush = true;

            // Create the binding.
            ExchangeService service;
            service = new ExchangeService(ExchangeVersion.Exchange2010);


            service.Credentials = new WebCredentials(args[1], args[2]);
            service.TraceListener = new TraceListener();
            service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

            string smtpAddress = args[3];
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

            string exchangeUrl;
            exchangeUrl = "https://";
            exchangeUrl += args[4];
            exchangeUrl += "/EWS/Exchange.asmx";


            service.Url = new Uri(exchangeUrl);


            Console.WriteLine("AutodiscoverURL: " + service.Url);

            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;

            service.TraceEnabled = true;



            FolderView folderView = new FolderView(5000);
            folderView.Traversal = FolderTraversal.Deep;

            FolderId rootFolderId = new FolderId(WellKnownFolderName.MsgFolderRoot);
           



            FindFoldersResults findFoldersResults = service.FindFolders(rootFolderId, folderView);
            if (findFoldersResults.Folders.Count > 0)
            {
                for (int i = 0; i < findFoldersResults.Folders.Count; i++)
                {
                    Folder fold = findFoldersResults.Folders[i];
                    Console.WriteLine("Folder:\t" + fold.DisplayName);
                    Console.WriteLine("Folder:\t" + fold.Id);

                    //Folder fold1;
                    //PropertySet p = new PropertySet(BasePropertySet.IdOnly, FolderSchema.DisplayName, FolderSchema.ChildFolderCount, FolderSchema.EffectiveRights, FolderSchema.FolderClass);

                    //fold1 = Folder.Bind(service, fold.Id, p);

                    if (fold.DisplayName.Contains(args[5]))//"[DMS]"))
                    {
                        fold.DisplayName = fold.DisplayName.Replace(',', ' ');
                        Log.WriteLine(String.Format("{0},{1}", fold.DisplayName, fold.Id.UniqueId));
                        Console.WriteLine(fold.DisplayName);
                        //Console.WriteLine(fold.Id);
                        Console.WriteLine("");
                    }
                }


            }



        }
        // ///Delete
        public static void AutoDiscover(string[] args)
        {
            if (args.Length < 4)
            {
                Console.WriteLine("Syntax: <Command> <ImpersonatorSMTP> <Password> <Enduser>");
                Console.WriteLine("Example: AutoDiscover ImpersonatorSMTPAddress@domain.com password user1@domain.com");
                return;
            }

            // Create the binding.
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010);


            service.Credentials = new WebCredentials(args[1], args[2]);
            service.TraceListener = new TraceListener();
            service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

            string smtpAddress = args[3];
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);


            try
            {
                service.AutodiscoverUrl(smtpAddress, RedirectionUrlValidationCallback);
            }
            catch (AutodiscoverRemoteException ex)
            {
                Console.WriteLine("Exception thrown: " + ex.Error.Message);
            }

            Console.WriteLine("AutodiscoverURL: " + service.Url);
        }


        public static void TestConnection(string[] args)
        {
            if (args.Length < 6)
            {
                Console.WriteLine("Syntax: <Command> <ImpersonatorSMTP> <password> <end user> <exchange server name> <IsExch2007>");
                Console.WriteLine("Example: TestConnection ImpersonatorSMTPAddress@dev.local password user1@dev.local xchange.dev.local false");
                return;
            }

            // Create the binding.
            ExchangeService service;
            if (args[5].ToUpper() == "TRUE")
                service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
            else
                service = new ExchangeService();


            service.Credentials = new WebCredentials(args[1], args[2]);
            service.TraceListener = new TraceListener();
            service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

            string smtpAddress = args[3];
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

            string exchangeUrl;
            exchangeUrl = "https://";
            exchangeUrl += args[4];
            exchangeUrl += "/EWS/Exchange.asmx";

            
            service.Url = new Uri(exchangeUrl);

            //try
            //{
            //    service.AutodiscoverUrl(smtpAddress, RedirectionUrlValidationCallback);
            //}
            //catch (AutodiscoverRemoteException ex)
            //{
            //    Console.WriteLine("Exception thrown: " + ex.Error.Message);
            //}

            Console.WriteLine("AutodiscoverURL: " + service.Url);

            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;

            service.TraceEnabled = true;

            PropertySet p = new PropertySet(BasePropertySet.IdOnly);

            Folder ewsFolder = Folder.Bind(service, WellKnownFolderName.Inbox, p);

        }

        public static void FindItem()
        {
             ExchangeService service;
            service = new ExchangeService(ExchangeVersion.Exchange2010);


            service.Credentials = new WebCredentials("admin@imanage.microsoftonline.com", "!wov2014");
            service.TraceListener = new TraceListener();
            service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

            string smtpAddress = "kenneth.lay@imanage.microsoftonline.com";
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

            string exchangeUrl;
            exchangeUrl = "https://";
            exchangeUrl += "ch1prd0410.outlook.com";
            exchangeUrl += "/EWS/Exchange.asmx";


            service.Url = new Uri(exchangeUrl);


            Console.WriteLine("AutodiscoverURL: " + service.Url);

            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;

            service.TraceEnabled = true;

            var view = new ItemView(100) { PropertySet = new PropertySet { EmailMessageSchema.Id, ItemSchema.Subject } };

            String searchstring = "Data retention - e-mailRepaired Message";
            SearchFilter.IsEqualTo filter = new SearchFilter.IsEqualTo(EmailMessageSchema.Subject, searchstring);
            FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, filter, view);
            Console.Write("IsEqualTo: Total email count with the specified search string in the subject: " + findResults.TotalCount);



            //var findResults = service.FindItems(WellKnownFolderName.Inbox, view);
            //var bindResults = service.BindToItems(findResults.Select(r => r.Id), new PropertySet { EmailMessageSchema.MimeContent});
            //foreach (GetItemResponse getItemResponse in bindResults)
            //{
            //    string sMimeCont;

            //    Item item = getItemResponse.Item;
            //    //s = item.MimeContent;
            //    sMimeCont = item.MimeContent.ToString();

            //}

           

           
            //ItemView view = new ItemView(50);

            //view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.MimeContent);

            //view.Traversal = ItemTraversal.Shallow;

            //FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox,  view);

            //PropertySet propSet = new PropertySet(BasePropertySet.IdOnly);


            //ServiceResponseCollection<GetItemResponse> response = service.BindToItems(findResults.Select(r => r.MimeContent), propSet);

            //foreach (GetItemResponse getItemResponse in response)
            //{
            //    string s;
                
            //    Item item = getItemResponse.Item;
            //    s = item.MimeContent;
               
            //}

            //// Process each item.
            //foreach (Item myItem in findResults.Items)
            //{
            //    if (myItem is EmailMessage)
            //    {
            //        Console.WriteLine((myItem as EmailMessage).Subject);
            //    }

            //    else if (myItem is MeetingRequest)
            //    {
            //        Console.WriteLine((myItem as MeetingRequest).Subject);
            //    }
            //    else
            //    {
            //        // Else handle other item types.
            //    }
            //}

        }

        // BindFolder
        public static void BindFolder(string[] args)
        {
            if (args.Length < 6)
            {
                Console.WriteLine("Syntax: <Command> <ImpersonatorSMTP> <password> <end user> <exchange server name> <Folder name>");
                Console.WriteLine("Example: BindFolder ImpersonatorSMTPAddress@dev.local password user1@dev.local xchange.dev.local MyFolder");
                return;
            }

            // Create the binding.
            ExchangeService service;
            service = new ExchangeService(ExchangeVersion.Exchange2010);


            service.Credentials = new WebCredentials(args[1], args[2]);
            service.TraceListener = new TraceListener();
            service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

            string smtpAddress = args[3];
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

            string exchangeUrl;
            exchangeUrl = "https://";
            exchangeUrl += args[4];
            exchangeUrl += "/EWS/Exchange.asmx";

            
            service.Url = new Uri(exchangeUrl);

            
            Console.WriteLine("AutodiscoverURL: " + service.Url);

            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;

            service.TraceEnabled = true;

           

            FolderView folderView = new FolderView(1000);
            folderView.Traversal = FolderTraversal.Shallow;
            
            FolderId rootFolderId = new FolderId(WellKnownFolderName.Inbox);
            SearchFilter.SearchFilterCollection searchFilterCollection = new SearchFilter.SearchFilterCollection();            
            SearchFilter searchFilter1 = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, args[5]);
            searchFilterCollection.Add(searchFilter1);

            

            FindFoldersResults findFoldersResults = service.FindFolders(rootFolderId, searchFilterCollection, folderView);
            if (findFoldersResults.Folders.Count > 0)
            {
                Folder fold = findFoldersResults.Folders[0];
                Console.WriteLine("Folder:\t" + fold.DisplayName);
                Console.WriteLine("Folder:\t" + fold.Id);

                Folder fold1;
                PropertySet p = new PropertySet(BasePropertySet.IdOnly, FolderSchema.DisplayName, FolderSchema.ChildFolderCount, FolderSchema.EffectiveRights, FolderSchema.FolderClass);

                fold1 = Folder.Bind(service, fold.Id, p);

                Console.WriteLine(fold.DisplayName);
                Console.WriteLine(fold.ChildFolderCount);
                Console.WriteLine(fold.EffectiveRights);
                Console.WriteLine(fold.FolderClass);
                
                
            }

            

        }

        public static String GetConvertedEWSID(ExchangeService esb, String sID, String strSMTPAdd)
        {
            String sEWSID = String.Empty;

            // Create a request to convert identifiers. 
            AlternateId objAltID = new AlternateId();
            objAltID.Format = IdFormat.HexEntryId;
            objAltID.Mailbox = strSMTPAdd;
            objAltID.UniqueId = sID;

            //Convert  PR_ENTRYID identifier format to an EWS identifier. 
            AlternateIdBase objAltIDBase = esb.ConvertId(objAltID, IdFormat.EwsId);
            if (null != objAltIDBase)
            {
                AlternateId objAltIDResp = (AlternateId)objAltIDBase;
                sEWSID = objAltIDResp.UniqueId;
            }
            return sEWSID;
        }

        public static void FindItemsInSearchFolder()
        {
            ExchangeService service;
            service = new ExchangeService(ExchangeVersion.Exchange2010);


            service.Credentials = new WebCredentials("admin@imanage.microsoftonline.com", "!Manage.2015");
            service.TraceListener = new TraceListener();
            service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

            string smtpAddress = "jsmith@imanage.microsoftonline.com";
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

            string exchangeUrl;
            exchangeUrl = "https://";
            exchangeUrl += "ch1prd0410.outlook.com";
            exchangeUrl += "/EWS/Exchange.asmx";


            service.Url = new Uri(exchangeUrl);


            Console.WriteLine("AutodiscoverURL: " + service.Url);

            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;

            service.TraceEnabled = true;

            FolderView folderView = new FolderView(10);
            folderView.PropertySet = new PropertySet(FolderSchema.DisplayName);

            //try
            {
                int iCnt = 0;
                FindFoldersResults findResults = service.FindFolders(WellKnownFolderName.SearchFolders, folderView);

                foreach (Folder folder in findResults.Folders)
                {
                    // You cannot request only search folders in 
                    // a FindFolders request, so other search folders might also be present.
                    if (folder is SearchFolder && folder.DisplayName.Equals("WCSE"))
                    {
                        Console.WriteLine("\"{0}\" folder found.", folder.DisplayName);

                        ItemView view = new ItemView(50);

                        // Identify the Subject properties to return.
                        // Indicate that the base property will be the item identifier
                        view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject);

                        // Set the traversal to shallow. (Shallow is the default option; other options are Associated and SoftDeleted.)
                        view.Traversal = ItemTraversal.Shallow;

                        
                        FindItemsResults<Item> findResults1;
                        do
                        {
                             // Send the request to search the Inbox and get the results.
                            findResults1 = folder.FindItems(view);

                            // Process each item.
                            foreach (Item myItem in findResults1.Items)
                            {
                                
                                if (myItem is EmailMessage)
                                {
                                    Console.WriteLine((myItem as EmailMessage).Subject);                                        
                                }
                            }
                            iCnt += findResults1.Items.Count();
                            view.Offset += 50;
                        } while (findResults1.MoreAvailable);
                       
                    }
                }
                Console.WriteLine("coutn - {0}", iCnt);
            }

        }

        public static void CreateFolder()
        {
            try
            {
                ExchangeService service;
                service = new ExchangeService(ExchangeVersion.Exchange2010);


                service.Credentials = new WebCredentials("admin@imanage.microsoftonline.com", "!Manage.2015");
                service.TraceListener = new TraceListener();
                service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

                string smtpAddress = "jsmith@imanage.microsoftonline.com";
                service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

                string exchangeUrl;
                exchangeUrl = "https://";
                exchangeUrl += "ch1prd0410.outlook.com";
                exchangeUrl += "/EWS/Exchange.asmx";


                service.Url = new Uri(exchangeUrl);


                Console.WriteLine("AutodiscoverURL: " + service.Url);

                ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;

                service.TraceEnabled = true;

                // Create a custom folder class.
                Folder folder = new Folder(service);
                folder.DisplayName = "CustomFolder1";//"iManage Failed Emails";
                

                // Create the folder as a child of the Inbox folder.
                folder.Save(WellKnownFolderName.MsgFolderRoot);



                string sEwsId = "AAMkAGFkZTM1MjY3LWZiYzAtNDA1ZC04NWI3LTA1ZWRlYzE2NjVjZABGAAAAAAAehyvl2c+VRaNBUFlASUlpBwA5Thqx2ogYS5z4GmODBiBuAAIHIHAaAAA5Thqx2ogYS5z4GmODBiBuAAIx7hnWAAA=";

                var bindResults = service.BindToItems(new[] { new ItemId(sEwsId) }, new PropertySet(BasePropertySet.IdOnly,
                                                            ItemSchema.Subject, ItemSchema.ItemClass));

                foreach (GetItemResponse getItemResponse in bindResults)
                {
                    string sSub;

                    Item item = getItemResponse.Item;
                    sSub = item.Subject;
                    sSub = item.ItemClass;
                    FolderId id = new FolderId(folder.Id.ToString());
                    item.Move(id);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: {0} : ", ex.Message);
            }

        }

        public static void CreateSearchFolder()
        {
            try
            {
                ExchangeService service;
                service = new ExchangeService(ExchangeVersion.Exchange2010);


                service.Credentials = new WebCredentials("admin@imanage.microsoftonline.com", "!Manage.2015");
                service.TraceListener = new TraceListener();
                service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

                string smtpAddress = "jsmith@imanage.microsoftonline.com";
                service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

                string exchangeUrl;
                exchangeUrl = "https://";
                exchangeUrl += "ch1prd0410.outlook.com";
                exchangeUrl += "/EWS/Exchange.asmx";


                service.Url = new Uri(exchangeUrl);


                Console.WriteLine("AutodiscoverURL: " + service.Url);

                ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;

                service.TraceEnabled = true;
                bool bDelete = false;
                if (bDelete)
                {
                    FolderView folderView = new FolderView(5000);
                    folderView.PropertySet = new PropertySet(FolderSchema.DisplayName);

                    FindFoldersResults findResults = service.FindFolders(WellKnownFolderName.MsgFolderRoot, folderView);

                    foreach (Folder folder in findResults.Folders)
                    {
                        Console.WriteLine("\"{0}\" folder .", folder.DisplayName);
                        // You cannot request only search folders in 
                        // a FindFolders request, so other folders might also be present.
                        if (folder is SearchFolder && folder.DisplayName.Equals("WCSE_FolderMappings"))
                        {
                            Console.WriteLine("\"{0}\" folder found.", folder.DisplayName);

                            folder.Delete(DeleteMode.HardDelete);

                            Console.WriteLine("\"{0}\" folder deleted.", folder.DisplayName);
                        }
                        
                    }
                }
                ////////////////////
                SearchFolder searchFolder1 = new SearchFolder(service);

                // Use the following search filter to get all mail in the Inbox with the word "extended" in the subject line.
                SearchFilter.ContainsSubstring searchCriteria =
                  new SearchFilter.ContainsSubstring(ItemSchema.Subject, "Queued");

                searchFolder1.SearchParameters.RootFolderIds.Add(WellKnownFolderName.Inbox);
                searchFolder1.SearchParameters.RootFolderIds.Add(WellKnownFolderName.Calendar);
                searchFolder1.SearchParameters.RootFolderIds.Add(WellKnownFolderName.SentItems);
                searchFolder1.SearchParameters.RootFolderIds.Add(WellKnownFolderName.DeletedItems);
                searchFolder1.SearchParameters.RootFolderIds.Add(WellKnownFolderName.Drafts);
                searchFolder1.SearchParameters.Traversal = SearchFolderTraversal.Deep;
                searchFolder1.SearchParameters.SearchFilter = searchCriteria;
                searchFolder1.DisplayName = "QueuedInboxCal";
                try
                {
                    searchFolder1.Save(WellKnownFolderName.SearchFolders);
                    Console.WriteLine(searchFolder1.DisplayName + " added.");
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error - " + e.Message);
                }

                ///////////////////////
                SearchFolder searchFolder = new SearchFolder(service);
                searchFolder.DisplayName = "WCSE2";
                SearchFilter.SearchFilterCollection searchAndFilterCollection =
                                                    new SearchFilter.SearchFilterCollection(LogicalOperator.And);

                SearchFilter.SearchFilterCollection searchOrFilterCollection =
                                        new SearchFilter.SearchFilterCollection(LogicalOperator.Or);

                ExtendedPropertyDefinition emailGuidProp = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders,
                                                                                                   "x-autn-guid", MapiPropertyType.String);


                searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.WorkSite.Ems.Queued"));
                //searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.WorkSite.Ems.Filed"));
                searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Schedule.Meeting.Request"));
                searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Appointment"));
                searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.EAS"));
                searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.SMIME"));
                searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Schedule.Meeting.Resp.Pos"));
                searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note"));
                searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.ABCD"));

                searchAndFilterCollection.Add(new SearchFilter.Exists(emailGuidProp));

                searchAndFilterCollection.Add(searchOrFilterCollection);

                searchFolder.SearchParameters.SearchFilter = searchAndFilterCollection;

                searchFolder.SearchParameters.RootFolderIds.Add(WellKnownFolderName.Root);
                searchFolder.SearchParameters.Traversal = SearchFolderTraversal.Deep;

                
                searchFolder.Save(WellKnownFolderName.SearchFolders);

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: {0} : ", ex.Message);
            }
            
        }
       // start
        public static void SearchBasedOnGuid(string[] args)
        {
            if (args.Length < 5)
            {
                Console.WriteLine("Syntax: <Command> <ImpersonatorSMTP> <password> <end user> <exchange server name> <Folder name>");
                Console.WriteLine("Example: SEARCH-GUID ImpersonatorSMTPAddress@dev.local password user1@dev.local xchange.dev.local");
                return;
            }

            // Create the binding.
            ExchangeService service;
            service = new ExchangeService(ExchangeVersion.Exchange2010);


            service.Credentials = new WebCredentials(args[1], args[2]);
            service.TraceListener = new TraceListener();
            service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

            string smtpAddress = args[3];
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

            string exchangeUrl;
            exchangeUrl = "https://";
            exchangeUrl += args[4];
            exchangeUrl += "/EWS/Exchange.asmx";


            service.Url = new Uri(exchangeUrl);


            Console.WriteLine("AutodiscoverURL: " + service.Url);

            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;

            service.TraceEnabled = true;

            
            

                //if (bProcess)

            string sEwsId = "00000000A385273B283ED211B34000A0C91E15DA01003B996EE1B942A047A4DA7CCE51CD5BC300016BF081450000";// args[5];

            sEwsId = GetConvertedEWSID(service, "00000000D5322A260E7FD011B31B00A0C91E15DA0700DC0767B594235A4BAA59BA7BC7AC64CB00188AAC802B0000DC0767B594235A4BAA59BA7BC7AC64CB00188AACE1810000",
                                        smtpAddress);
            sEwsId = "AAMkADI5M2M4Zjg0LWNjODYtNGVmYy04ZmJiLTYwYjVmNmE1MmUxMQAuAAAAAAC5avaeLRx+RLLL+ESyKSkTAQA7mW7huUKgR6TafM5RzVvDAAHw7oDIAAA=";

            ExtendedPropertyDefinition emailGuidProp1 = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders,
                                                                                                    "x-autn-guid", MapiPropertyType.String);

            ExtendedPropertyDefinition filingStatus = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                                    "FilingStatus", MapiPropertyType.String);

            var bindResults = service.BindToItems(new[] { new ItemId(sEwsId) }, new PropertySet(BasePropertySet.IdOnly,
                                                            ItemSchema.Subject, ItemSchema.ItemClass, emailGuidProp1));
            foreach (GetItemResponse getItemResponse in bindResults)
            {
                string sSub;

                Item item = getItemResponse.Item;
                sSub = item.Subject;
                sSub = item.ItemClass;
                //sMimeCont = item.MimeContent.ToString();

                foreach (ExtendedProperty extendedProperty in item.ExtendedProperties)
                {
                    if (extendedProperty.PropertyDefinition.Name == filingStatus.Name &&
                            extendedProperty.PropertyDefinition.PropertySetId == filingStatus.PropertySetId)
                    {
                        item.RemoveExtendedProperty(filingStatus);
                        break;
                    }


                }

                foreach (ExtendedProperty extendedProperty in item.ExtendedProperties)
                {
                    if (extendedProperty.PropertyDefinition.Name == emailGuidProp1.Name &&
                            extendedProperty.PropertyDefinition.PropertySetId == emailGuidProp1.PropertySetId)
                    {
                        item.RemoveExtendedProperty(filingStatus);
                        break;
                    }
                }

                

            }
            if (sEwsId.Length > 0)
            {
                ///////////
                AlternateId objAltID = new AlternateId();
                objAltID.Format = IdFormat.HexEntryId;
                objAltID.Mailbox = smtpAddress;
                objAltID.UniqueId = sEwsId;

                //Convert  PR_ENTRYID identifier format to an EWS identifier. 
                AlternateIdBase objAltIDBase = service.ConvertId(objAltID, IdFormat.EwsId);
                if (null != objAltIDBase)
                {
                    AlternateId objAltIDResp = (AlternateId)objAltIDBase;
                    sEwsId = objAltIDResp.UniqueId;
                }
                ///////////////
                Folder fld;
                FolderId id = new FolderId(sEwsId);

                fld = Folder.Bind(service, id);
                Console.WriteLine("Folder Name: " + fld.DisplayName);


                ExtendedPropertyDefinition emailGuidProp = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders,
                                                                                    "x-autn-guid", MapiPropertyType.String);
                

                SearchFilter.SearchFilterCollection searchFilterCollection =
                                        new SearchFilter.SearchFilterCollection(LogicalOperator.And);
                searchFilterCollection.Add(new SearchFilter.IsEqualTo(emailGuidProp, "E6054A91-643D-4ED5-A82C-8CF7DD017D46"));//"79F353FC-0F6F-445E-8D60-AE6AD3AF7559"));
                searchFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.WorkSite.Ems.Filed"));
                



                FindItemsResults<Item> findResults;
                //Collection<EmailMessage> 
                do
                {
                    ItemView view = new ItemView(50);

                    // Identify the Subject properties to return.
                    // Indicate that the base property will be the item identifier
                    view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, emailGuidProp);

                    // Set the traversal to shallow. (Shallow is the default option; other options are Associated and SoftDeleted.)
                    view.Traversal = ItemTraversal.Shallow;


                    // Send the request to search the Inbox and get the results.
                    findResults = service.FindItems(id, searchFilterCollection, view);


                    int extendedPropertyindex = 0;
                    bool bUpdate = false;

                    // Process each item.
                    foreach (Item myItem in findResults.Items)
                    {
                        bUpdate = true;
                        extendedPropertyindex = 0;

                        if (myItem is EmailMessage)
                        {
                            Console.WriteLine((myItem as EmailMessage).Subject);                                        
                        }
                    }
                    
                    view.Offset += 50;
                } while (findResults.MoreAvailable);

            }
        
    
            

            #region Commented
            //string sEwsId;
            //sEwsId = GetConvertedEWSID(service, "000000001E872BE5D9CF9545A3415059404949690100394E1AB1DA88184B9CF81A638306206E0000C40D0F1D0000",
            //                            smtpAddress);

            //Folder fld;
            //FolderId id = new FolderId(sEwsId);

            //fld = Folder.Bind(service, id);
            //Console.WriteLine("Folder Name: " + fld.DisplayName);
            // Add a search filter that searches on the body or subject.
            ////List<SearchFilter> searchFilterCollection = new List<SearchFilter>();
            ////searchFilterCollection.Add(new SearchFilter.ContainsSubstring(ItemSchema.Subject, "Test"));
            ////searchFilterCollection.Add(new SearchFilter.ContainsSubstring(ItemSchema.Body, "homecoming"));

            //ExtendedPropertyDefinition emailGuidProp = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders,
            //                                                                    "x-autn-guid", MapiPropertyType.String);

            //ItemView view = new ItemView(50);

            //// Identify the Subject properties to return.
            //// Indicate that the base property will be the item identifier
            //view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, emailGuidProp);

            //// Set the traversal to shallow. (Shallow is the default option; other options are Associated and SoftDeleted.)
            //view.Traversal = ItemTraversal.Shallow;
            //FindItemsResults<Item> findResults;
            //do
            //{
            //    // Send the request to search the Inbox and get the results.
            //    findResults = service.FindItems(id, /*searchFilter,*/ view);

            //    int extendedPropertyindex = 0;

            //    // Process each item.
            //    foreach (Item myItem in findResults.Items)
            //    {
            //        extendedPropertyindex = 0;

            //        if (myItem is EmailMessage)
            //        {
            //            foreach (ExtendedProperty extendedProperty in myItem.ExtendedProperties)
            //            {
            //                if (extendedProperty.PropertyDefinition.Name == emailGuidProp.Name &&
            //                        extendedProperty.PropertyDefinition.PropertySetId == emailGuidProp.PropertySetId)
            //                {
            //                    myItem.RemoveExtendedProperty(emailGuidProp);
            //                    break;
            //                }

            //                extendedPropertyindex++;
            //            }
            //            myItem.ItemClass = "IPM.Note";

            //            myItem.Update(ConflictResolutionMode.AlwaysOverwrite);

            //            Console.WriteLine((myItem as EmailMessage).Subject);
            //        }
            //    }
            //    view.Offset += 50;
            //} while (findResults.MoreAvailable);
            #endregion
        } 

        // end
        public static void CustomMapiProvCleanup(string[] args)
        {
            if (args.Length < 5)
            {
                Console.WriteLine("Syntax: <Command> <ImpersonatorSMTP> <password> <end user> <exchange server name> <Folder name>");
                Console.WriteLine("Example: PROPCLEANUP ImpersonatorSMTPAddress@dev.local password user1@dev.local xchange.dev.local");
                return;
            }

            // Create the binding.
            ExchangeService service;
            service = new ExchangeService(ExchangeVersion.Exchange2010);


            service.Credentials = new WebCredentials(args[1], args[2]);
            service.TraceListener = new TraceListener();
            service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

            string smtpAddress = args[3];
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

            string exchangeUrl;
            exchangeUrl = "https://";
            exchangeUrl += args[4];
            exchangeUrl += "/EWS/Exchange.asmx";


            service.Url = new Uri(exchangeUrl);


            Console.WriteLine("AutodiscoverURL: " + service.Url);

            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;

            service.TraceEnabled = true;

            if (!File.Exists("EWSConfig.txt"))
            {
                Console.WriteLine("EWSConfig.txt doesn't exist");
                return;
            }

            System.IO.StreamReader file = new System.IO.StreamReader("EWSConfig.txt");
            string line;
            bool bProcess = false;
            while ((line = file.ReadLine()) != null)
            {

                if (line == "[ResetCustomMAPIProp]")
                {
                    bProcess = true;
                    continue;
                }
                else
                {
                    if (line.Contains("["))
                        bProcess = false;
                }

                if (bProcess)
                {
                    string sEwsId;
                    if (line.Length > 0)
                    {
                        sEwsId = GetConvertedEWSID(service, line, smtpAddress);

                        if (sEwsId.Length > 0)
                        {
                            Folder fld;
                            FolderId id = new FolderId(sEwsId);

                            fld = Folder.Bind(service, id);
                            Console.WriteLine("Folder Name: " + fld.DisplayName);


                            ExtendedPropertyDefinition emailGuidProp = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders,
                                                                                                "x-autn-guid", MapiPropertyType.String);


                            SearchFilter.SearchFilterCollection searchFilterCollection =
                                                    new SearchFilter.SearchFilterCollection(LogicalOperator.Or);

                            searchFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.WorkSite.Ems.Filed"));
                            searchFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.WorkSite.Ems.Error"));


                           
                            FindItemsResults<Item> findResults;
                            //Collection<EmailMessage> 
                            do
                            {
                                ItemView view = new ItemView(50);

                                // Identify the Subject properties to return.
                                // Indicate that the base property will be the item identifier
                                view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, emailGuidProp);

                                // Set the traversal to shallow. (Shallow is the default option; other options are Associated and SoftDeleted.)
                                view.Traversal = ItemTraversal.Shallow;

                                
                                // Send the request to search the Inbox and get the results.
                                findResults = service.FindItems(id, searchFilterCollection, view);

                                
                                int extendedPropertyindex = 0;
                                bool bUpdate = false;

                                // Process each item.
                                foreach (Item myItem in findResults.Items)
                                {
                                    bUpdate = true;
                                    extendedPropertyindex = 0;

                                    if (myItem is EmailMessage)
                                    {
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

                                       // myItem.Update(ConflictResolutionMode.AlwaysOverwrite);

                                        Console.WriteLine((myItem as EmailMessage).Subject);
                                    }
                                }
                                if (bUpdate)
                                    service.UpdateItems(findResults, id, ConflictResolutionMode.AlwaysOverwrite, MessageDisposition.SaveOnly, null);

                                view.Offset += 50;
                            } while (findResults.MoreAvailable);
                            
                        }
                    }
                }
            }

            #region Commented
            //string sEwsId;
            //sEwsId = GetConvertedEWSID(service, "000000001E872BE5D9CF9545A3415059404949690100394E1AB1DA88184B9CF81A638306206E0000C40D0F1D0000",
            //                            smtpAddress);

            //Folder fld;
            //FolderId id = new FolderId(sEwsId);
            
            //fld = Folder.Bind(service, id);
            //Console.WriteLine("Folder Name: " + fld.DisplayName);
            // Add a search filter that searches on the body or subject.
            ////List<SearchFilter> searchFilterCollection = new List<SearchFilter>();
            ////searchFilterCollection.Add(new SearchFilter.ContainsSubstring(ItemSchema.Subject, "Test"));
            ////searchFilterCollection.Add(new SearchFilter.ContainsSubstring(ItemSchema.Body, "homecoming"));

            //ExtendedPropertyDefinition emailGuidProp = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders,
            //                                                                    "x-autn-guid", MapiPropertyType.String);

            //ItemView view = new ItemView(50);

            //// Identify the Subject properties to return.
            //// Indicate that the base property will be the item identifier
            //view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, emailGuidProp);

            //// Set the traversal to shallow. (Shallow is the default option; other options are Associated and SoftDeleted.)
            //view.Traversal = ItemTraversal.Shallow;
            //FindItemsResults<Item> findResults;
            //do
            //{
            //    // Send the request to search the Inbox and get the results.
            //    findResults = service.FindItems(id, /*searchFilter,*/ view);

            //    int extendedPropertyindex = 0;

            //    // Process each item.
            //    foreach (Item myItem in findResults.Items)
            //    {
            //        extendedPropertyindex = 0;

            //        if (myItem is EmailMessage)
            //        {
            //            foreach (ExtendedProperty extendedProperty in myItem.ExtendedProperties)
            //            {
            //                if (extendedProperty.PropertyDefinition.Name == emailGuidProp.Name &&
            //                        extendedProperty.PropertyDefinition.PropertySetId == emailGuidProp.PropertySetId)
            //                {
            //                    myItem.RemoveExtendedProperty(emailGuidProp);
            //                    break;
            //                }

            //                extendedPropertyindex++;
            //            }
            //            myItem.ItemClass = "IPM.Note";

            //            myItem.Update(ConflictResolutionMode.AlwaysOverwrite);

            //            Console.WriteLine((myItem as EmailMessage).Subject);
            //        }
            //    }
            //    view.Offset += 50;
            //} while (findResults.MoreAvailable);
            #endregion
        }        

        static void Main(string[] args)
        {
            //CreateFolder();
            //CreateSearchFolder();
            //FindItemsInSearchFolder();
            //SplitCSV oSplitCSV = new SplitCSV();
            //oSplitCSV.SplitCSVForLinkedFolders("WO-33282-1.csv");
//            FindLinkedFolders(args);
            //SearchBasedFolderId(args);
            //SearchBasedOnGuid(args);
            //SplitCSV oSplitCSV = new SplitCSV();
            //oSplitCSV.SplitCSVFileWithValidEntryIdOrGuid("wo-25464-query-output.csv");//FB-5-19.csv");
            //Console.WriteLine("DONE");
            //return;
            if (args.Length > 0)
            {
                for (int i = 0; i < args.Length; i++ )
                {
                    Console.WriteLine(args[i]);
                }

                if (args[0].ToUpper() == "AUTODISCOVER")
                {
                    AutoDiscover(args); 
                }
                else if (args[0].ToUpper() == "TESTCONNECTION")
                {
                    TestConnection(args);
                }
                else if (args[0].ToUpper() == "BINDFOLDER")
                {
                    BindFolder(args);
                }
                else if (args[0].ToUpper() == "PROPCLEANUP")
                {
                    CustomMapiProvCleanup(args);
                }
                else if (args[0].ToUpper() == "FIXPLAINTEXT-STEP2")
                {
                    FixPlainTextStep2 command = new FixPlainTextStep2();
                    command.Execute(args);
                }
                else if (args[0].ToUpper() == "SCAN-FOLDERS")
                {
                    ResetEmails command = new ResetEmails();
                    command.Execute(args);
                }
                else if (args[0].ToUpper() == "SCAN-EMAIL")
                {
                    ResetEmails command = new ResetEmails();
                    command.ScanEmailWithEntryId(args);
                }
                else if (args[0].ToUpper() == "SCAN-OUTLOOK-FOLDERS")
                {
                    ResetEmails command = new ResetEmails();
                    command.ExecuteScanOutlookFolders(args);
                }
                else if (args[0].ToUpper() == "GET-LINKED-FOLDERS")
                {
                    FindLinkedFolders(args);
                }
                else if (args[0].ToUpper() == "SCAN-LINKED-FOLDERS")
                {
                    ResetEmails command = new ResetEmails();
                    //command.ScanLinkedFolders(args);
                    command.ScanLinkedFoldersEx(args);
                }
                else if (args[0].ToUpper() == "SCAN-FILED-QUEUED-EMAILS")
                {
                    ResetEmails command = new ResetEmails();
                    //command.ScanLinkedFolders(args);
                    command.ScanLinkedFoldersForFilingStatusFiledMsgClsQueued(args);
                }
                Console.WriteLine("Done");
            }
            else
            {
                Console.WriteLine("Invalid Command");
                Console.WriteLine("Available commands 1) Autodiscover 2) TestConnection 3) BindFolder 4) PropCleanup 5) FixPlainText-Step2 6)SCAN-FOLDERS");
                Console.WriteLine("for more info on a particular command type EWSTestApp CommandName");
            }
        }
    }
}
            

