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

            FolderId rootFolderId = new FolderId(WellKnownFolderName.Root);
           



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
                        //FolderId f = new FolderId(fold.ParentFolderId.UniqueId);
                        //Folder fld = Folder.Bind(service, f);
                        //Console.WriteLine(fld.DisplayName);
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
            //ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            ExchangeService service = new ExchangeService();

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

        public static void SetAndGetConfig()
        {
            ExchangeService service;
            service = new ExchangeService(ExchangeVersion.Exchange2010);


            service.Credentials = new WebCredentials("ewsuser@exdev2016.local", "!manage6");
            service.TraceListener = new TraceListener();
            service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

            string smtpAddress = "user1@exdev2016.local";
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

            string exchangeUrl;
            exchangeUrl = "https://";
            exchangeUrl += "10.192.211.238";
            exchangeUrl += "/EWS/Exchange.asmx";


            service.Url = new Uri(exchangeUrl);


            Console.WriteLine("AutodiscoverURL: " + service.Url);

            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;

            service.TraceEnabled = true;

            // Create the user configuration object.
            UserConfiguration config = new UserConfiguration(service);
            
            // Add user configuration data to the XmlData and BinaryData properties.
           // config.Name = "NewProfessionalClient";

            // Name and save the user configuration object on the Inbox folder.
            // This results in a call to EWS.
            //config.Save("NewProfessionalClient", WellKnownFolderName.Inbox);

            //Get

            // Bind to a user configuration object. This results in a call to EWS.
            UserConfiguration usrConfig = UserConfiguration.Bind(service,
                                                                 "NewProfessionalClient",
                                                                 WellKnownFolderName.Inbox,
                                                                 UserConfigurationProperties.All);
            // Display the returned property values.
            Console.WriteLine("User Config Identifier: " + usrConfig.ItemId.UniqueId);
            Console.WriteLine("Name: " + usrConfig.Name);
           // Console.WriteLine("XmlData: " + Encoding.UTF8.GetString(usrConfig.XmlData));
           // Console.WriteLine("BinaryData: " + Encoding.UTF8.GetString(usrConfig.BinaryData));
        }
        public static void FindItem()
        {
             ExchangeService service;
            service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);


            service.Credentials = new WebCredentials("ewsuser@exdev2016.local", "!manage6");
            //service.Credentials = new WebCredentials("wcssvc@Support.cg", "!nterw0ven");
            
            service.TraceListener = new TraceListener();
            service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

            string smtpAddress = "user1@exdev2016.local";
            //string smtpAddress = "bnichting@support.cg";
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

            string exchangeUrl;
            exchangeUrl = "https://";
            exchangeUrl += "10.192.211.238";
            //exchangeUrl += "10.192.224.248";
;            exchangeUrl += "/EWS/Exchange.asmx";

           
            service.Url = new Uri(exchangeUrl);


            Console.WriteLine("AutodiscoverURL: " + service.Url);

            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;

            service.TraceEnabled = true;

            
            GetUserRetentionPolicyTagsResponse RetResp = service.GetUserRetentionPolicyTags();
            foreach (RetentionPolicyTag myItem in RetResp.RetentionPolicyTags)
            {
                Guid g = myItem.RetentionId;
                String s = myItem.DisplayName;
                s = myItem.Description;
                RetentionActionType type = myItem.RetentionAction;
                int i = myItem.RetentionPeriod;
                ElcFolderType fType = myItem.Type;
                
            }
            ExtendedPropertyDefinition emailGuidProp = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders,
                                                                                              "x-wsguid", MapiPropertyType.String);

            ExtendedPropertyDefinition filingStatus = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                        "FilingLocation", MapiPropertyType.String);
            
            var view = new ItemView(100) { PropertySet = new PropertySet { EmailMessageSchema.Id, ItemSchema.Subject, ItemSchema.ItemClass, ItemSchema.Categories, emailGuidProp, filingStatus } };
           
           // String searchstring = "Test S&N - 113 [IWOV-WS_DB_94.FID14]";
            //SearchFilter.IsEqualTo filter = new SearchFilter.IsEqualTo(EmailMessageSchema.Subject, searchstring);
            SearchFilter.IsEqualTo filter = new SearchFilter.IsEqualTo(EmailMessageSchema.Categories, "Queued");
            FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, filter, view);
            Console.Write("IsEqualTo: Total email count with the specified search string in the subject: " + findResults.TotalCount);
            int extendedPropertyindex = 0;
            foreach (Item myItem in findResults.Items)
            {
                extendedPropertyindex = 0;
                if (myItem is EmailMessage)
                {
                   // myItem.Categories.Clear();
                    myItem.Categories.Add("Queued");
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
                    myItem.Update(ConflictResolutionMode.AlwaysOverwrite);
                }
            }

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
                PropertySet p = new PropertySet(BasePropertySet.IdOnly, FolderSchema.TotalCount, FolderSchema.DisplayName, FolderSchema.ChildFolderCount, FolderSchema.EffectiveRights, FolderSchema.FolderClass);

                fold1 = Folder.Bind(service, fold.Id, p);

                Console.WriteLine(fold.TotalCount);
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

        public static void SyncSearchFolder()
        {
            ExchangeService service;
            service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);


            //service.Credentials = new WebCredentials("admin@imanage.microsoftonline.com", "!Manage.2015");
            service.Credentials = new WebCredentials("WCSadmin@wcsdev.net", "!manage5");
            service.TraceListener = new TraceListener();
            service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

            string smtpAddress = "thopulou@wcsdev.net";
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

            string exchangeUrl;
            exchangeUrl = "https://";
            exchangeUrl += "10.8.210.113";
            exchangeUrl += "/EWS/Exchange.asmx";


            service.Url = new Uri(exchangeUrl);


            Console.WriteLine("AutodiscoverURL: " + service.Url);

            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;

            service.TraceEnabled = true;
            //string s1 = "/o=ExDev2016/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=bf14deac35b843a3990bd94fbf81da5a-User1";
            //service.ResolveName(s1,ResolveNameSearchLocation.DirectoryOnly, false);
           
           //String EWSId = "AAMkADlhYzczZmVkLTRjZWUtNDE4My1iMjFlLWVlZmFjZGVjMDgzMgBGAAAAAACP5kXq3pBqQ4PTuAsvacgqBwDuPLFEltGYS465LpM2aReMAAAAAAEMAADuPLFEltGYS465LpM2aReMAAAmOfy0AAA=";
           // String s = "0000000038A1BB1005E5101AA1BB08002B2A56C20000454D534D44422E444C4C00000000000000001B55FA20AA6611CD9BC800AA002FC45A0C0000007573657231406578646576323031362E6C6F63616C002F6F3D4578446576323031362F6F753D45786368616E67652041646D696E6973747261746976652047726F7570202846594449424F484632335350444C54292F636E3D526563697069656E74732F636E3D62663134646561633335623834336133393930626439346662663831646135612D557365723100E94632F43E00000002000000100000007500730065007200310040006500780064006500760032003000310036002E006C006F00630061006C0000000000";
           // AlternateId aiRequest1 = new AlternateId();

           // aiRequest1.UniqueId = s;
           // aiRequest1.Mailbox = "user1@exdev2016.local";
           // aiRequest1.Format = IdFormat.HexEntryId;
           // AlternateId aiResultsStore1 = (AlternateId)service.ConvertId(aiRequest1, IdFormat.StoreId);
           // Console.WriteLine(aiResultsStore1.Mailbox);

           // AlternateId aiRequest = new AlternateId();
           
           // aiRequest.UniqueId = EWSId;
           // aiRequest.Mailbox = "user67686@exdev2016.local";
           // aiRequest.Format = IdFormat.EwsId;
           // AlternateId aiResultsStore = (AlternateId)service.ConvertId(aiRequest, IdFormat.StoreId);
           // Console.WriteLine(aiResultsStore.Mailbox);

           // FolderView folderView = new FolderView(1000);
           // folderView.PropertySet = new PropertySet(FolderSchema.DisplayName);

            ////////////////////////////////
            //ItemView iv = new ItemView(1);
            //ExtendedPropertyDefinition PR_Search_Key = new ExtendedPropertyDefinition(0x300B, MapiPropertyType.Binary);
            ////PropertySet psProperSet = new PropertySet();
            ////psProperSet.Add(PR_Search_Key);
            ////iv.PropertySet = psProperSet;
            ////FindItemsResults<Item> fiItems = service.FindItems(WellKnownFolderName.Inbox, iv);
            ////Item SourceItem = fiItems.Items[0];
            ////SourceItem.Move(WellKnownFolderName.JunkEmail);
            ////Console.WriteLine(SourceItem.ExtendedProperties[0].Value.ToString());
            ////foreach (ExtendedProperty extendedProperty in SourceItem.ExtendedProperties)
            ////{
            ////    Console.WriteLine(extendedProperty.PropertyDefinition.Name.ToString());

            ////}
            //SearchFilter sfSearchFilter = new SearchFilter.IsEqualTo(PR_Search_Key, "K9oqT1UJZEmeYUUgOSEHJA==");//"Vi6DyQ2Ou0G/CSgAo1n1nA==");//Convert.ToBase64String((byte[])SourceItem.ExtendedProperties[0].Value));
            //FindItemsResults<Item> mvMoveItems = service.FindItems(WellKnownFolderName.JunkEmail, sfSearchFilter, iv);
            //Item MovedItem = mvMoveItems.Items[0];
            //////////////////////////////////
            //try
            {
                //////////////Subscription///////////////
                //string _ListenURi = "http://10.5.8.195:36728/ews-notify";//10.5.8.195
                //PushSubscription _subscription1 = null;
                //_subscription1 = service.SubscribeToPushNotificationsOnAllFolders(new Uri(_ListenURi), 1, "", EventType.Moved);
                //return;
                //////////////Subscription///////////////

                string cSyncState = "";
                int iCnt = 0;
                FolderView folderView = new FolderView(1);
                folderView.PropertySet = new PropertySet(FolderSchema.DisplayName);
                
                //SearchFilter searchFoldFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "WCSE_FolderMappings");
                //FindFoldersResults findResults = service.FindFolders(WellKnownFolderName.Root, folderView);


                //foreach (Folder folder in findResults.Folders)
                {
                    //Console.WriteLine("\"{0}\" folder found.", folder.DisplayName);

                   // if (folder is SearchFolder && folder.DisplayName.Equals("WCSE_FolderMappings"))
                    //if (folder.DisplayName.Equals("WCSE_FolderMappings"))//WCSE_SFMailboxSync"))
                    {
                        
                        //Console.WriteLine("\"{0}\" folder found.", folder.DisplayName);
                        
                        //SearchFolder f = SearchFolder.Bind(service, folder.Id);
                        //SearchFilter filter = f.SearchParameters.SearchFilter;
                       // string s = filter.ToString();
                        //Console.WriteLine(s);
                       ItemView view = new ItemView(1000);
                       //view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.ParentFolderId, ItemSchema.LastModifiedTime);
                        view.Traversal = ItemTraversal.Shallow;

                        ExtendedPropertyDefinition emailGuidProp1 = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders,
                                                                                                    "x-autn-guid", MapiPropertyType.String);

                        //ExtendedPropertyDefinition eDef = new ExtendedPropertyDefinition(0x300B, MapiPropertyType.Binary);
                        ExtendedPropertyDefinition eDef = new ExtendedPropertyDefinition(12299, MapiPropertyType.Binary);
                        view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, eDef, emailGuidProp1);

                        ExtendedPropertyDefinition prSub = new ExtendedPropertyDefinition(0x0037, MapiPropertyType.String);
                        //view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, prSub);
                        view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, emailGuidProp1, prSub);

                       // PropertySet prop = BasePropertySet.IdOnly;
                        //prop.Add(eDef); 

                        String sDt = "2016-03-21T17:14:31Z";

                       // SearchFilter.IsGreaterThanOrEqualTo searchCriteria1 =
                                  //  new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.LastModifiedTime, sDt);
                        //searchFolder.SearchParameters.SearchFilter = new SearchFilter.IsEqualTo(EmailMessageSchema.Sender, smtpAddress);
                        SearchFilter.SearchFilterCollection filterAllItemsFolder = new SearchFilter.SearchFilterCollection(LogicalOperator.And);
                        filterAllItemsFolder.Add(new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.LastModifiedTime, sDt));
                        filterAllItemsFolder.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.Sender, "thopulou@wcsdev.net")); 
                        //SearchFilter.IsEqualTo searchCriteria1 =
                        //                new SearchFilter.IsEqualTo(prSub, "Feb 18 - 003");
                        //SearchFilter.IsEqualTo searchCriteria1 =
                        //            new SearchFilter.IsEqualTo(emailGuidProp1, "1A260F3D-3B0C-4CC6-8337-FD51530D818D");
                        //string searchval = "7602A4544A677747BE6A1E3203E36E7B";
                        //byte[] array = Encoding.ASCII.GetBytes(searchval);
                        //SearchFilter searchCriteria1 = new SearchFilter.IsEqualTo(eDef, Convert.ToBase64String((byte[])SourceItem.ExtendedProperties[0].Value));
                        //SearchFilter.IsEqualTo searchCriteria1 =
                        //            new SearchFilter.IsEqualTo(eDef, "A0DA33B14532544683BE08ADFF9956B0");

                        FindItemsResults<Item> findResults1 = service.FindItems(WellKnownFolderName.Inbox, filterAllItemsFolder, view);

                        foreach (Item myItem in findResults1.Items)
                        {
                            if (myItem is EmailMessage)
                            {
                                Console.WriteLine((myItem as EmailMessage).Subject);
                                //Console.WriteLine((myItem as EmailMessage).LastModifiedTime.ToString());
                                //Console.WriteLine((myItem as EmailMessage).ParentFolderId.UniqueId);

                                //Folder f1 = Folder.Bind(service, (myItem as EmailMessage).ParentFolderId.UniqueId);
                                //Console.WriteLine(f1.DisplayName);
                               
                                foreach (ExtendedProperty extendedProperty in myItem.ExtendedProperties)
                                {
                                    Console.WriteLine(extendedProperty.PropertyDefinition.Name.ToString());

                                }

                                //Byte[] PropVal;
                                //String HexSearchKey;
                                //if (myItem.TryGetProperty(eDef, out PropVal))
                                //{                               
                                //    HexSearchKey = BitConverter.ToString(PropVal).Replace("-", "");
                                //    Console.WriteLine(HexSearchKey);
                                //}
                                
                            }
                        }
                    }
                }
            }
        }
        public static void delete()
        {

        }


        private static bool GetAllItemsFolder(ref ExchangeService service, ref Folder AllItems)
        {
            bool bRet = false;
            //Log1.WriteLine("> GetAllItemsFolder");

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

            //Log1.WriteLine("< GetAllItemsFolder");
            return bRet;
        }

        public static void FindItemsInEntireMailbox(string[] args)
        {
             StreamWriter Log1 = new StreamWriter("EWSTestAppLog.txt", true);
             Log1.AutoFlush = true;
            int iTotalEmailsReset = 0;
            int iCount = 0;
            int iTotalCnt = 0;
            int iSkippedForThisFolder = 0;

            if (args.Length < 7)
            {
                Console.WriteLine("Syntax: <Command> <ImpersonatorSMTP> <password> <end user> <exchange server name> <csvFile> <QueuedMsg>");
                Console.WriteLine("Example: SCAN-IN-MAILBOX ewsuser@exdev2016.local password user1@exdev2016.local 10.192.211.238 EmRequests.csv True");
                return;
            }


            do
            {

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
                    Log1.WriteLine("Exchange server field is blank");
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


                String EMRequest = args[5];
                String sQueueOnly = args[6];

                if (!File.Exists(EMRequest))
                {
                    Log1.WriteLine("File doesn't exist - {0}", EMRequest);
                    Console.WriteLine("File doesn't exist - {0}", EMRequest);
                    break;
                }

                Folder AllItems = null;
                bool bRet = GetAllItemsFolder(ref service, ref AllItems);
                if (null == AllItems)
                {
                    bRet = false;
                    break;
                }

                if (!File.Exists("MessageClasses.txt"))
                {
                    Log1.WriteLine("File doesn't exist - MessageClasses.txt");
                    Console.WriteLine("File doesn't exist - MessageClasses.txt");
                    break;
                }

                SearchFilter.SearchFilterCollection searchOrFilterCollection =
                                                new SearchFilter.SearchFilterCollection(LogicalOperator.Or);

                System.IO.StreamReader fileMsgClass = new System.IO.StreamReader("MessageClasses.txt");

                string lineMsgCls;
                while ((lineMsgCls = fileMsgClass.ReadLine()) != null)
                {
                    lineMsgCls.Trim();
                    if (lineMsgCls.Length > 1)
                        searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, lineMsgCls));
                }


                System.IO.StreamReader file = new System.IO.StreamReader(EMRequest);

                Dictionary<String, String> oEntryIds = null;
                oEntryIds = new Dictionary<String, String>();


                string line;
                string emailguid;
                string entryid;

                while ((line = file.ReadLine()) != null)
                {
                    line.Trim();
                    String[] Tokens = line.Split(",".ToCharArray());
                    if (Tokens.Length < 2)
                    {
                        Log1.WriteLine("Invalid data in {0}", EMRequest);
                        break;
                    }
                    emailguid = Tokens[1].ToString();
                    entryid = Tokens[0].ToString();

                 

                    //////////////
                    SearchFilter.SearchFilterCollection searchFilterCollection =
                                                new SearchFilter.SearchFilterCollection(LogicalOperator.And);

                    ExtendedPropertyDefinition filingStatus = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                                        "FilingStatus", MapiPropertyType.String);

                    ExtendedPropertyDefinition emailGuidProp = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders,
                                                                                               "x-autn-guid", MapiPropertyType.String);

                    if (sQueueOnly.ToUpper() == "TRUE")
                        searchFilterCollection.Add(new SearchFilter.IsEqualTo(filingStatus, "Queued"));
                    else
                        searchFilterCollection.Add(new SearchFilter.IsEqualTo(filingStatus, "Filed"));

                    searchFilterCollection.Add(new SearchFilter.IsEqualTo(emailGuidProp, emailguid));

                    ////////////////



                  
                    FindItemsResults<Item> findResults;
                    ItemView view = new ItemView(100, 0, OffsetBasePoint.Beginning);

                    // Identify the Subject properties to return.
                    // Indicate that the base property will be the item identifier
                    view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.ItemClass, filingStatus);

                    // Set the traversal to shallow. (Shallow is the default option; other options are Associated and SoftDeleted.)
                    view.Traversal = ItemTraversal.Shallow;

                    do
                    {

                        findResults = service.FindItems(AllItems.Id, searchFilterCollection, view);
                        if (findResults.Count() == 0)
                        {
                            PropertySet PropsToFetch = null;
                            PropsToFetch = new PropertySet();
                            PropsToFetch.Add(ItemSchema.Id);
                            PropsToFetch.Add(ItemSchema.Subject);

                            AlternateId objAltID = new AlternateId();
                            objAltID.Format = IdFormat.HexEntryId;
                            objAltID.Mailbox = smtpAddress;
                            objAltID.UniqueId = entryid;

                            //Convert  PR_ENTRYID identifier format to an EWS identifier. 
                            try
                            {
                                AlternateIdBase objAltIDBase = service.ConvertId(objAltID, IdFormat.EwsId);
                                AlternateId objAltIDResp = (AlternateId)objAltIDBase;

                                ServiceResponseCollection<GetItemResponse> bindResults = null;

                                //bindResults = service.BindToItems(findResults.Select(r => r.Id), PropsToFetch);
                                ItemId[] itemIDList = { objAltIDResp.UniqueId };

                                bindResults = service.BindToItems(itemIDList, PropsToFetch);
                                Item OutlookItem = null;
                                for (int nIter = 0; nIter < bindResults.Count; nIter++)
                                {
                                    OutlookItem = bindResults[nIter].Item;
                                    Log1.WriteLine("Email Exist: Subject - {0}", OutlookItem.Subject);
                                    Log1.WriteLine("Email Exist: Id: Id - {0}", OutlookItem.Id);
                                    iCount++;
                                }
                            }
                            catch(Exception)
                            { }
                        }
                        else
                        {
                            // Process each item.
                            foreach (Item myItem in findResults.Items)
                            {
                                if (myItem is EmailMessage)
                                {
                                    Log1.WriteLine("Found Email : Subject - {0}", (myItem as EmailMessage).Subject);
                                    Log1.WriteLine("Found Email : EWSId - {0}", myItem.Id.UniqueId);
                                    Log1.WriteLine("");
                                    Console.WriteLine("Found Email : {0}", (myItem as EmailMessage).Subject);
                                    iTotalEmailsReset++;
                                }
                            }
                            //iCount += findResults.Items.Count();
                            iTotalCnt += findResults.Items.Count();
                        }
                        
                        //if (sReportMode.ToUpper() == "TRUE")
                        //    view.Offset += 100;
                        //else
                        //    view.Offset = iSkippedForThisFolder;
                    } while (findResults.MoreAvailable);

                }
                

            } while (false);

            Log1.WriteLine("");
            Log1.WriteLine("Total emails reset using email guid {0} ", iTotalEmailsReset);
            Log1.WriteLine("Total emails reset using EntryId {0} ", iCount);
            //Log1.WriteLine("");
            //Log1.WriteLine("Total emails skipped (request exist in em_req) for {0} on Folder: {1} - {2} - {3}", smtpAddress, FoldName, FoldEwsId, iSkippedQueuedEMails);
            //Log1.WriteLine("");
        }

        public static void FindItemsInSearchFolder()
        {
            ExchangeService service;
            service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);


            //service.Credentials = new WebCredentials("admin@imanage.microsoftonline.com", "!Manage.2015");
            service.Credentials = new WebCredentials("ewsuser@exdev2016.local", "!manage6");
            service.TraceListener = new TraceListener();
            service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

            string smtpAddress = "user1@exdev2016.local";
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

            string exchangeUrl;
            exchangeUrl = "https://";
            exchangeUrl += "10.192.211.238";
            exchangeUrl += "/EWS/Exchange.asmx";


            service.Url = new Uri(exchangeUrl);


            Console.WriteLine("AutodiscoverURL: " + service.Url);

            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;

            service.TraceEnabled = true;

            FolderView folderView = new FolderView(1000);
            folderView.PropertySet = new PropertySet(FolderSchema.DisplayName);

            //try
            {
                int iCnt = 0;
                FindFoldersResults findResults = service.FindFolders(WellKnownFolderName.MsgFolderRoot, folderView);

                foreach (Folder folder in findResults.Folders)
                {
                    Console.WriteLine("\"{0}\" folder found.", folder.DisplayName);
                    // You cannot request only search folders in 
                    // a FindFolders request, so other search folders might also be present.
                    if (folder is SearchFolder && folder.DisplayName.Equals("WCSE_SFMailboxSync"))
                    //if (folder.DisplayName.Equals("WCSE_FolderMappings"))//WCSE_SFMailboxSync"))
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
                                PropertySet pset = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject,
                                                        EmailMessageSchema.ParentFolderId,
                                                        EmailMessageSchema.From, EmailMessageSchema.InternetMessageId, 
                                                        EmailMessageSchema.ToRecipients);
                                Item i1 = Item.Bind(service, myItem.Id, pset);

                                //fold1 = Folder.Bind(service, fold.Id, p);
                                Folder f1 = Folder.Bind(service, i1.ParentFolderId);
                                Console.WriteLine(f1.DisplayName);
                                //Console.WriteLine((i1 as EmailMessage).From);
                                //Console.WriteLine((i1 as EmailMessage).InternetMessageId);//ToRecipients.Contains("jsmith@imanage.microsoftonline.com"));
                                //Console.WriteLine((i1 as EmailMessage).ToRecipients.Contains("jsmith123@imanage.microsoftonline.com"));
                                
                                //if (myItem is EmailMessage)
                                //{
                                //    Console.WriteLine((myItem as EmailMessage).From);//Subject);                                        
                                //}
                            }
                            iCnt += findResults1.Items.Count();
                            view.Offset += 50;
                        } while (findResults1.MoreAvailable);
                       
                    }
                }
                Console.WriteLine("coutn - {0}", iCnt);
            }

        }


        public static void CreateHiddenSearchFolder()
        {
            try
            {
                ExchangeService service;
                service = new ExchangeService(ExchangeVersion.Exchange2010);


                service.Credentials = new WebCredentials("ewsuser@exdev2016.local", "!manage6");
                service.TraceListener = new TraceListener();
                service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

                string smtpAddress = "user1@exdev2016.local";
                service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

                string exchangeUrl;
                exchangeUrl = "https://";
                exchangeUrl += "10.192.211.238";
                exchangeUrl += "/EWS/Exchange.asmx";


                service.Url = new Uri(exchangeUrl);


                Console.WriteLine("AutodiscoverURL: " + service.Url);

                ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;

                service.TraceEnabled = true;

                // Create a custom folder class.
                //SearchFolder folder = new SearchFolder(service);

                //folder.DisplayName = "WCSE2_Hid1";

                //folder.SearchParameters.SearchFilter = new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.WorkSite.Ems.Filed");

                //folder.SearchParameters.RootFolderIds.Add(WellKnownFolderName.MsgFolderRoot);
                //folder.SearchParameters.Traversal = SearchFolderTraversal.Deep;


                //folder.Save(WellKnownFolderName.Inbox);

                //folder.DisplayName = "Hidden Folder 2";
                
                //// Create the folder as a child of the Inbox folder.
                //folder.Save(WellKnownFolderName.Inbox);

                //ExtendedPropertyDefinition isHiddenProp = new ExtendedPropertyDefinition(0x10f4, MapiPropertyType.Boolean);
                //PropertySet propSet = new PropertySet(isHiddenProp);

                //// Bind to a folder and retrieve the PidTagAttributeHidden property.
                //Folder folder1 = Folder.Bind(service, folder.Id, propSet);

                //// Set the PidTagAttributeHidden property to true.
                //folder1.SetExtendedProperty(isHiddenProp, true);

                // Save the changes.
                //folder1.Update();

                //string sEwsId = "AQMkADdiYTJmZgEwLTBmNzAtNDkxNy1iZDcyLTU2YzIyAjYyZTZlAEYAAAO6I5nos0G+Tpe6VTf4HSgvBwAxbf8ZIHVHTIHMLg3c37BfAAACAQwAAAAxbf8ZIHVHTIHMLg3c37BfAAAA3p73fwAAAA==";
                //"AQMkADdiYTJmZgEwLTBmNzAtNDkxNy1iZDcyLTU2YzIyAjYyZTZlAEYAAAO6I5nos0G+Tpe6VTf4HSgvBwAxbf8ZIHVHTIHMLg3c37BfAAACAQkAAAAxbf8ZIHVHTIHMLg3c37BfAAACKwwAAAA=";
                //EmailMessage beforeMessage = EmailMessage.Bind(service, new ItemId(sEwsId), propSet);

                //NameResolutionCollection resolvedNames = service.ResolveName("ExdevTestUsers");
                //// Output the list of candidates.
                //foreach (NameResolution nameRes in resolvedNames)
                //{
                //    Console.WriteLine("Contact name: " + nameRes.Mailbox.Name);
                //    Console.WriteLine("Contact e-mail address: " + nameRes.Mailbox.Address);
                //    Console.WriteLine("Mailbox type: " + nameRes.Mailbox.MailboxType);
                //}
                //ExpandGroupResults myGroupMembers = service.ExpandGroup("ExdevTestUsers");

                //foreach (EmailAddress address in myGroupMembers.Members)
                //{
                //    Console.WriteLine("Email Address: {0}", address);
                //}

                
                ////folder.Move(WellKnownFolderName.SearchFolders);

                //return;

                ///////////////////////
                SearchFolder searchFolder = new SearchFolder(service);
                searchFolder.DisplayName = "WCSE2_Hid1";
                
                //searchFolder.SearchParameters.SearchFilter = new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.WorkSite.Ems.Queued");
                searchFolder.SearchParameters.SearchFilter = new SearchFilter.IsEqualTo(EmailMessageSchema.Sender, smtpAddress);
                //searchFolder.SearchParameters.SearchFilter = new SearchFilter.IsEqualTo(EmailMessageSchema.Sender, eName);

                searchFolder.SearchParameters.RootFolderIds.Add(WellKnownFolderName.MsgFolderRoot);
                searchFolder.SearchParameters.Traversal = SearchFolderTraversal.Deep;


                searchFolder.Save(WellKnownFolderName.SearchFolders);

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: {0} : ", ex.Message);
            }

        }

        public static void MoveItem()
        {
            try
            {
                
                ExchangeService service;
                service = new ExchangeService(ExchangeVersion.Exchange2010);


                service.Credentials = new WebCredentials("ewsuser@exdev2016.local", "!manage6");
                //service.Credentials = new WebCredentials("user6@exdev2016.local", "!manage9");
                service.TraceListener = new TraceListener();
                service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

                string smtpAddress = "user1@exdev2016.local";
                service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

                string exchangeUrl;
                exchangeUrl = "https://";
                exchangeUrl += "10.192.211.238";
                exchangeUrl += "/EWS/Exchange.asmx";


                service.Url = new Uri(exchangeUrl);


                Console.WriteLine("AutodiscoverURL: " + service.Url);

                ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;
                
                
                ///
                Folder f = new Folder(service);
                ExtendedPropertyDefinition isHiddenProp2 = new ExtendedPropertyDefinition(0x10f4, MapiPropertyType.Boolean);
                PropertySet propSet2 = new PropertySet(isHiddenProp2);
                f.SetExtendedProperty(isHiddenProp2, true);
                
                f.DisplayName = "CreateHidden";
                f.Save(WellKnownFolderName.SentItems);

                ///Make it hidden
                ///
                string sEwsFId = "AQMkADdiYTJmZgEwLTBmNzAtNDkxNy1iZDcyLTU2YzIyAjYyZTZlAC4AAAO6I5nos0G+Tpe6VTf4HSgvAQAxbf8ZIHVHTIHMLg3c37BfAAAAtqj4iwAAAA==";
                FolderId fid = new FolderId(sEwsFId);
                FolderView fView = new FolderView(1000);
                service.FindFolders(fid, fView);

                // Create an extended property definition for the PidTagAttributeHidden property.
                ExtendedPropertyDefinition isHiddenProp = new ExtendedPropertyDefinition(0x10f4, MapiPropertyType.Boolean);
                PropertySet propSet1 = new PropertySet(isHiddenProp);

                // Bind to a folder and retrieve the PidTagAttributeHidden property.
                Folder folder = Folder.Bind(service, fid, propSet1);

                // Set the PidTagAttributeHidden property to true.
                folder.SetExtendedProperty(isHiddenProp, true);

                // Save the changes.
                folder.Update();
                //////////////

                //service.MoveItems();
                ////
                PropertySet propSet = new PropertySet(BasePropertySet.IdOnly, EmailMessageSchema.Subject, EmailMessageSchema.ParentFolderId);

                
                // Bind to the existing item by using the ItemId.
                // This method call results in a GetItem call to EWS.
                string sEwsId = "AQMkADdiYTJmZgEwLTBmNzAtNDkxNy1iZDcyLTU2YzIyAjYyZTZlAEYAAAO6I5nos0G+Tpe6VTf4HSgvBwAxbf8ZIHVHTIHMLg3c37BfAAACAQwAAAAxbf8ZIHVHTIHMLg3c37BfAAAAsgEuJwAAAA==";
                //"AQMkADdiYTJmZgEwLTBmNzAtNDkxNy1iZDcyLTU2YzIyAjYyZTZlAEYAAAO6I5nos0G+Tpe6VTf4HSgvBwAxbf8ZIHVHTIHMLg3c37BfAAACAQkAAAAxbf8ZIHVHTIHMLg3c37BfAAACKwwAAAA=";
                EmailMessage beforeMessage = EmailMessage.Bind(service, new ItemId(sEwsId), propSet);

                Console.WriteLine(beforeMessage.Subject);

                // Move the specified mail to the JunkEmail folder and store the returned item.
                Item item1 = beforeMessage.Move(WellKnownFolderName.Inbox);
                
                // Check that the item was moved by binding to the moved email message 
                // and retrieving the new ParentFolderId.
                // This method call results in a GetItem call to EWS.
                //EmailMessage movedMessage = EmailMessage.Bind(service, item.Id, propSet);

               // Console.WriteLine("An email message with the subject '" + beforeMessage.Subject + "' has been moved from the '" + before
                
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
                ///////////////////////////////////////////////////////
               // string exchangeUrl1;
               // exchangeUrl1 = "https://";
               // exchangeUrl1 += "10.192.211.238";
               // exchangeUrl1 += "/EWS/Exchange.asmx";
               // EWSUtil.EWSConnection ews1 = new EWSUtil.EWSConnection("user7@exdev2016.local", true, "ewsuser", "!manage6", "exdev2016", exchangeUrl1);
               // //ews1.esb.CookieContainer ccont = new System.Net.CookieContainer();
                
               // EWSUtil.List<EWSUtil.EWS.DelegateUserResponseMessageType> list = ews1.getDeletgates("user7@exdev2016.local");
                
               //// EWSUtil.List<EWSUtil.EWS.DelegateUserResponseMessageType> list = ews1.esb.GetDelegate("user1@exdev2016.local");

               // foreach (EWSUtil.EWS.DelegateUserResponseMessageType res in list)
               // {
               //     if (res.DelegateUser.ReceiveCopiesOfMeetingMessages.Equals(true))
               //         Console.WriteLine("ReceiveCopiesOfMeetingMessages");

               //     if (res.DelegateUser.ReceiveCopiesOfMeetingMessagesSpecified.Equals(true))
               //         Console.WriteLine("ReceiveCopiesOfMeetingMessagesSpecified");

               //     if (res.DelegateUser.ViewPrivateItemsSpecified.Equals(true))
               //         Console.WriteLine("ViewPrivateItemsSpecified");

               //     if (res.DelegateUser.ViewPrivateItems.Equals(true))
               //         Console.WriteLine("ViewPrivateItems");
               // }

                //// Bind to the service by using the primary e-mail address credentials.
                //ExchangeService service1 = new ExchangeService(ExchangeVersion.Exchange2010);
                //service1.Credentials = new WebCredentials("user9@exdev2016.local", "!manage8");
                //service1.TraceListener = new TraceListener();
                //service1.TraceFlags = TraceFlags.All;
                ////service1.Credentials = new WebCredentials("user1@exdev2016.local", "!manage9");
                //service1.Url = new Uri("https://10.192.211.238/ews/exchange.asmx");
                
                //ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;
                //// Create a mailbox object that represents the primary user.
                //Mailbox mailbox1 = new Mailbox("user9@exdev2016.local");
                ////Mailbox mailbox1 = new Mailbox("user1@exdev2016.local");

                //// Call the GetDelegates method to get the delegates of the primary user.
                //DelegateInformation result1 = service1.GetDelegates(mailbox1, true);
                
                ///////////////////////////////////////////////////////

                ExchangeService service;
                service = new ExchangeService(ExchangeVersion.Exchange2010);


                service.Credentials = new WebCredentials("ewsuser@exdev2016.local", "!manage6");
                //service.Credentials = new WebCredentials("user7@exdev2016.local", "!manage8");
                service.TraceListener = new TraceListener();
                service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

                string smtpAddress = "user1@exdev2016.local";
                service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

                string exchangeUrl;
                exchangeUrl = "https://";
                exchangeUrl += "10.192.211.238";
                exchangeUrl += "/EWS/Exchange.asmx";


                service.Url = new Uri(exchangeUrl);


                Console.WriteLine("AutodiscoverURL: " + service.Url);

                ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;

                

                //Mailbox mailbox = new Mailbox(smtpAddress);
                //DelegateInformation result = service.GetDelegates(mailbox, true);
                service.TraceEnabled = true;
                bool bDelete = false;
                if (bDelete)
                {
                    FolderView folderView = new FolderView(1000);
                    folderView.PropertySet = new PropertySet(FolderSchema.DisplayName);

                    FindFoldersResults findResults = service.FindFolders(WellKnownFolderName.Root, folderView);

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
                  new SearchFilter.ContainsSubstring(ItemSchema.Subject, "Filed");

                //searchFolder1.SearchParameters.RootFolderIds.Add(WellKnownFolderName.Inbox);
                //searchFolder1.SearchParameters.RootFolderIds.Add(WellKnownFolderName.Calendar);
                //searchFolder1.SearchParameters.RootFolderIds.Add(WellKnownFolderName.SentItems);
                //searchFolder1.SearchParameters.RootFolderIds.Add(WellKnownFolderName.DeletedItems);
                searchFolder1.SearchParameters.RootFolderIds.Add(WellKnownFolderName.Root);
                searchFolder1.SearchParameters.Traversal = SearchFolderTraversal.Deep;
                searchFolder1.SearchParameters.SearchFilter = searchCriteria;
                searchFolder1.DisplayName = "QueuedInboxCal101";
                try
                {
                    searchFolder1.Save(WellKnownFolderName.Root);
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


        public static void DeleteFolder(string[] args)
        {
            if (args.Length < 7)
            {
                Console.WriteLine("Syntax: <Command> <ImpersonatorSMTP> <password> <exchange server name> <file name> <ReportMode>");
                Console.WriteLine("Example: DEL-IMANAGE-FOLDER ewsuser@exdev.com password ExchangeServer DeleteiManageFolder.txt True");
                return;
            }

           // try
            {

                System.IO.StreamReader fileMsgClass = new System.IO.StreamReader(args[5]);

                string lineMsgCls;
                while ((lineMsgCls = fileMsgClass.ReadLine()) != null)
                {
                    ExchangeService service;
                    service = new ExchangeService(ExchangeVersion.Exchange2010);


                    service.Credentials = new WebCredentials(args[1], args[2]);
                    //service.Credentials = new WebCredentials("user7@exdev2016.local", "!manage8");
                    service.TraceListener = new TraceListener();
                    service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

                    string smtpAddress = lineMsgCls;
                    service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

                    string exchangeUrl;
                    exchangeUrl = "https://";
                    exchangeUrl += args[3];
                    exchangeUrl += "/EWS/Exchange.asmx";


                    service.Url = new Uri(exchangeUrl);


                    Console.WriteLine("AutodiscoverURL: " + service.Url);

                    ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;


                    service.TraceEnabled = true;


                    FolderView folderView = new FolderView(10);
                    folderView.PropertySet = new PropertySet(FolderSchema.DisplayName);

                //    SearchFilter.SearchFilterCollection searchFilterCollection;
                //searchFilterCollection.Add(new SearchFilter.IsEqualTo(emailGuidProp, "E6054A91-643D-4ED5-A82C-8CF7DD017D46"));//"79F353FC-0F6F-445E-8D60-AE6AD3AF7559"));
                //searchFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.WorkSite.Ems.Filed"));

                    FindFoldersResults findResults = service.FindFolders(WellKnownFolderName.Root, 
                                                                            new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "FileName"), 
                                                                            folderView);

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

                    service.Credentials = null;
                    service.ImpersonatedUserId = null;
                    service.TraceListener = null;
                    service = null;


                }

            }

        }


        static void Main(string[] args)
        {
            
            //CreateHiddenSearchFolder();
            //CreateSearchFolder();

            //WorkSiteUtility im = new WorkSiteUtility();
            //im.Login("10.192.211.228", "ewsuser", "mhdocs");
            //im.test();
            //return;
           // MoveItem();
            //SetAndGetConfig();
            //SyncSearchFolder();
            //FindItem();
            //return;
            //return;
           // FindItemsInEntireMailbox(args);
            //CreateSearchFolder();
            //FindItemsInSearchFolder();
            //SyncSearchFolder();
            
           // SplitCSV oSplitCSV = new SplitCSV();
            //oSplitCSV.SplitCSVForLinkedFolders(args);//"WO-33282-1.csv");
            //return;
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
                    //command.ScanLinkedFolders(args);//old                    
                    command.ScanLinkedFoldersEx(args);
                }
                else if (args[0].ToUpper() == "SCAN-MAILBOX-RESET-QUEUED-EMAILS")
                {
                    ResetEmails command = new ResetEmails();
                    //command.ScanLinkedFolders(args);//old
                    command.ScanMailboxToResetEmails(args);
                    //command.ScanLinkedFoldersEx(args);
                }
                else if (args[0].ToUpper() == "SCAN-FILED-QUEUED-EMAILS")
                {
                    ResetEmails command = new ResetEmails();
                    //command.ScanLinkedFolders(args);
                    command.ScanLinkedFoldersForFilingStatusFiledMsgClsQueued(args);
                }
                else if (args[0].ToUpper() == "DELETE-SEARCH-FOLDER")
                {
                    ResetEmails command = new ResetEmails();
                    command.DeleteSearchFolder(args);
                    //command.RecoverItemsFromDumpster(args);
                }
                    //scan for search folder which has the same name as WCSE search folder
                else if (args[0].ToUpper() == "SCAN-WCSE-SEARCH-FOLDER")
                {
                    ScanForSearchFolder command = new ScanForSearchFolder();
                    command.ScanSearchFolder(args);
                    //command.RecoverItemsFromDumpster(args);
                }
                else if (args[0].ToUpper() == "SCAN-SENT-ITEM-FOLDER")
                {
                    ScanSentItemFolderForSendAndFile command = new ScanSentItemFolderForSendAndFile();
                    command.ScanSentItemFolder(args);
                    //command.RecoverItemsFromDumpster(args);
                }
                else if (args[0].ToUpper() == "IMPORT-MSG-TO-FOLDER")
                {
                    ImportMsgToFolder command = new ImportMsgToFolder();
                    command.ImportMsgToSpecifiedFolder(args);
                }
                // Decode Quoted Printable and Base64 and write the decodedFrom feild to CSV file
                else if (args[0].ToUpper() == "DECODE-FROM-FIELD")
                {
                    string inputfile = args[1];
                    DecodeFromField command = new DecodeFromField();
                    command.ReadCsvFileToDecode(inputfile);
                }
                else if (args[0].ToUpper() == "SCAN-RECOVERY-FOLDER")
                {
                    ResetEmails command = new ResetEmails();
                    //command.DeleteSearchFolder(args);
                    command.RecoverItemsFromDumpster(args);
                }
                else if (args[0].ToUpper() == "SCAN-RETENTION_POLICY_TAGS")
                {
                    ResetEmails command = new ResetEmails();
                    command.ScanItemsForRetentionPolicyTags(args);
                }
                else if (args[0].ToUpper() == "UPDATE_SEARCH_CRITERIA")
                {
                    ResetEmails command = new ResetEmails();
                    command.UpdateSearchCriteria(args);
                }
                else if (args[0].ToUpper() == "SPLITXML")
                {
                    SplitCSV oSplitCSV = new SplitCSV();
                    oSplitCSV.SplitCSVForLinkedFolders(args);//"WO-33282-1.csv");       
                }
                else if (args[0].ToUpper() == "GEN_BAT_ENTIRE_MB")
                {
                    SplitCSV oSplitCSV = new SplitCSV();
                    oSplitCSV.GenerateBatchFileToResetEmailsOnEntireMailbox(args);//"WO-33282-1.csv");       
                }
                else if (args[0].ToUpper() == "UPDATE_MSG_CLS_BASED_ON_FILING_STATUS") // June 2017
                {
                    WorkSiteUtility workUtility = new WorkSiteUtility();
                    ExchangeUtilityFunctions oUtility = new ExchangeUtilityFunctions(ref workUtility);
                    oUtility.CreateLogFile(1, "EWSTestAppLog.txt");
                    oUtility.UpdateMsgClsBasedOnFilingStatus(args);
                }
                else if (args[0].ToUpper() == "GET_ALL_FOLDER_MAPPINGS")
                {
                    WorkSiteUtility workUtility = new WorkSiteUtility();
                    if (!workUtility.Login(args[1], args[2], args[3]))
                    {
                        Console.WriteLine("Login to WorkServer failed: {0}, {1}, {2}", args[1], args[2], args[3]);
                        Console.WriteLine("Login to WorkServer failed");
                    }
                    else
                    {
                        ExchangeUtilityFunctions oUtility = new ExchangeUtilityFunctions(ref workUtility);
                        oUtility.CreateLogFile(1, "EWSTestAppLog.txt");
                        oUtility.CreateLogFile(3, "WCSUtilReport-MappedFolders.csv");
                        if (!oUtility.Initialize(args[2], args[3]))
                        {
                            Console.WriteLine("Invalid entries in DatabaseConfig.txt");
                        }
                        else
                            oUtility.ScanAllFolderMappings(args);

                        oUtility.Cleanup();
                    }
                }
                //delete filed items from mapped folder when unchecking 'leave message in outlook after filing'

                else if (args[0].ToUpper() == "DELETE_ITEMS_AFTER_MAPPING")
                {
                    WorkSiteUtility workUtility = new WorkSiteUtility();
                    if (!workUtility.Login(args[1], args[2], args[3]))
                    {
                        Console.WriteLine("Login to WorkServer failed: {0}, {1}, {2}", args[1], args[2], args[3]);
                        Console.WriteLine("Login to WorkServer failed");
                    }
                    else
                    {
                        DeleteItemFromMappedFolders oUtility = new DeleteItemFromMappedFolders(ref workUtility);
                        oUtility.CreateLogFile(1, "EWSTestAppLog.txt");
                        oUtility.CreateLogFile(3, "WCSUtilReport-MappedFolders.csv");
                        if (!oUtility.Initialize(args[2], args[3]))
                        {
                            Console.WriteLine("Invalid entries in DatabaseConfig.txt");
                        }
                        else
                            oUtility.DeleteFiledItemsFromMappedFolder(args);

                        oUtility.Cleanup();
                    }
                }
                else if (args[0].ToUpper() == "SCAN_SEARCH_FOLDER")
                {
                    if (args.Length < 13)
                    {
                        Console.WriteLine("Syntax: <Command> <WorkServer> <NRTAdmin> <password> <ImpersonatorSMTP> <password> <exchange server name> <ExchangeVersion> <SearchFolderParent> <SearchFolder> <FilingStatus> <CountOnly>");
                        //Console.WriteLine("Example: SCAN_SEARCH_FOLDER WorkSite NRTAdmin password ImpersonatorSMTPAddress@dev.local password exchangeServer Exchange2010 2 1 Queued False");
                        Console.WriteLine("Example: SCAN_SEARCH_FOLDER WorkSite NRTAdmin password ImpersonatorSMTPAddress@dev.local password exchangeServer Exchange2010 2 1 Queued False True");
                        return;
                    }
                    string sReportMode = args[12].ToUpper();
                    while (true)
                    {
                        // SCAN_SEARCH_FOLDER 10.192.211.228 ewsuser mhdocs ewsuser@exdev2016.local !manage6 10.192.211.238 Exchange2010_SP1 2 1 Queued False False
                        WorkSiteUtility workUtility = new WorkSiteUtility();

                        if (!workUtility.Login(args[1], args[2], args[3]))//"10.192.211.228", "ewsuser", "mhdocs");
                        {
                            Console.WriteLine("Login to WorkServer failed: {0}, {1}, {2}", args[1], args[2], args[3]);
                            Console.WriteLine("Login to WorkServer failed");
                        }
                        else
                        {
                            ExchangeUtilityFunctions oUtility = new ExchangeUtilityFunctions(ref workUtility);
                            oUtility.CreateLogFile(1, "EWSTestAppLog.txt");

                            if (sReportMode == "TRUE")
                                oUtility.CreateLogFile(2, "WCSUtilReport_ReportMode.csv");
                            else
                                oUtility.CreateLogFile(2, "WCSUtilReport_FixMode.csv");


                            if (!oUtility.Initialize(args[2], args[3]))
                            {
                                Console.WriteLine("Invalid entries in DatabaseConfig.txt");

                            }
                            else
                                oUtility.ScanSearchFolderForEmails(args);

                            oUtility.Cleanup();
                        }
                        workUtility = null;

                        if (sReportMode == "TRUE")
                            break;

                        sReportMode = "TRUE";
                    }
                }
                else if (args[0].ToUpper() == "CREATE_UNFILED_SEARCH_FOLDER")
                {
                    WorkSiteUtility workUtility = new WorkSiteUtility();
                    ExchangeUtilityFunctions oUtility = new ExchangeUtilityFunctions(ref workUtility);
                    oUtility.CreateLogFile(1, "EWSTestAppLog.txt");
                    oUtility.CreateUnfiledSearchFolder(args);
                    oUtility.Cleanup();
                }
                else if (args[0].ToUpper() == "GET-MF-NOT-QUEUED-EMAILS")
                {
                    WorkSiteUtility workUtility = new WorkSiteUtility();
                    ExchangeUtilityFunctions oUtility = new ExchangeUtilityFunctions(ref workUtility);
                    oUtility.CreateLogFile(1, "EWSTestAppLog.txt");
                    oUtility.CreateLogFile(4, "WCSUtilReportMappedFolders.csv");
                    oUtility.GetNotQueuedEmailsFromMappedFolder(args);
                    oUtility.Cleanup();
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
            

 