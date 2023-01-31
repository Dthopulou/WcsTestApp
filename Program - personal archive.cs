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

namespace EWSTestApp
{
    class Program
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



        public static bool CertificateValidationCallback(Object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }

        static bool RedirectionUrlValidationCallback(String redirectionUrl)
        {

            return true;
        }

        public static void GetItemsFromFolders(string[] args)
        {
            ExchangeService service;
            service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);


            service.Credentials = new WebCredentials(args[1], args[2]);//"admin@imanage.microsoftonline.com", "!wov2014");
            service.TraceListener = new TraceListener();
            service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

            string smtpAddress = "jsmith@imanage.microsoftonline.com";
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

            string exchangeUrl;
            exchangeUrl = "https://";
            exchangeUrl += args[4];// "ch1prd0410.outlook.com";
            exchangeUrl += "/EWS/Exchange.asmx";


            service.Url = new Uri(exchangeUrl);


            Console.WriteLine("AutodiscoverURL: " + service.Url);

            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;

            service.TraceEnabled = true;

            Folder fld;

            Console.WriteLine("Id: " + args[5]);
            FolderId id = new FolderId(args[5]);

            fld = Folder.Bind(service, id);
            Console.WriteLine("");
            Console.WriteLine("Folder Name: " + fld.DisplayName);

            ////////////////////

            if (fld.DisplayName.Length > 0)
            {
                var view = new ItemView(100) { PropertySet = new PropertySet { EmailMessageSchema.Id, ItemSchema.Subject, ItemSchema.Id } };
                view.Traversal = ItemTraversal.Shallow;

                //String searchstring = "RE: new bits";// "test march 20 - 001";
                SearchFilter.IsLessThan filter = new SearchFilter.IsLessThan(EmailMessageSchema.DateTimeReceived, args[6]);//"2015-03-16T14:15:50Z");

                FindItemsResults<Item> findResults = service.FindItems(fld.Id, filter, view);
                Console.Write("Total number of emails " + findResults.TotalCount);
                Console.WriteLine("");
                foreach (Item myItem in findResults.Items)
                {
                    if (myItem is EmailMessage)
                    {
                        Console.WriteLine((myItem as EmailMessage).Subject);
                        Console.WriteLine((myItem as EmailMessage).Id);
                        

                        AlternateId objAltID = new AlternateId();
                        objAltID.Format = IdFormat.EwsId;
                        objAltID.Mailbox = args[3];// "jsmith@imanage.microsoftonline.com";
                        objAltID.UniqueId = (myItem as EmailMessage).Id.ToString();

                        //Convert  PR_ENTRYID identifier format to an EWS identifier. 
                        AlternateIdBase objAltIDBase = service.ConvertId(objAltID, IdFormat.HexEntryId);
                        AlternateId objAltIDResp = (AlternateId)objAltIDBase;
                        Console.WriteLine(objAltIDResp.UniqueId);
                    }
                }
            }

        }


        public static void GetEntryIdFromEWSId(string sID)
        {
            ExchangeService service;
            service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);


            service.Credentials = new WebCredentials("ewsuser@dev.local", "!manage32");
            service.TraceListener = new TraceListener();
            service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

            string smtpAddress = "vinod@dev.local";
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

            string exchangeUrl;
            exchangeUrl = "https://";
            exchangeUrl += "xchange.dev.local";
            exchangeUrl += "/EWS/Exchange.asmx";


            service.Url = new Uri(exchangeUrl);


            Console.WriteLine("AutodiscoverURL: " + service.Url);

            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;

            service.TraceEnabled = true;

            // Create a request to convert identifiers. 
            AlternateId objAltID = new AlternateId();
            objAltID.Format = IdFormat.EwsId;
            objAltID.Mailbox = "vinod@dev.local";
            objAltID.IsArchive = true;
            objAltID.UniqueId = sID;

            //Convert  PR_ENTRYID identifier format to an EWS identifier. 
            AlternateIdBase objAltIDBase = service.ConvertId(objAltID, IdFormat.HexEntryId);
            AlternateId objAltIDResp = (AlternateId)objAltIDBase;

            Console.WriteLine(objAltIDResp.UniqueId);
            //sEwsId1 = GetConvertedEntryID(service, "AAMkAGFkZTM1MjY3LWZiYzAtNDA1ZC04NWI3LTA1ZWRlYzE2NjVjZABGAAAAAAAehyvl2c+VRaNBUFlASUlpBwA5Thqx2ogYS5z4GmODBiBuAAF/PIMQAAA5Thqx2ogYS5z4GmODBiBuAAGi4SjGAAA=",
            //                          smtpAddress);
        }

        public static void GetEWSIdFromEntryId(string sID)
        {
            ExchangeService service;
            service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);


            service.Credentials = new WebCredentials("ewsuser@dev.local", "!manage32");
            service.TraceListener = new TraceListener();
            service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

            string smtpAddress = "vinod@dev.local";
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

            string exchangeUrl;
            exchangeUrl = "https://";
            exchangeUrl += "xchange.dev.local";
            exchangeUrl += "/EWS/Exchange.asmx";


            service.Url = new Uri(exchangeUrl);


            Console.WriteLine("AutodiscoverURL: " + service.Url);

            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;

            service.TraceEnabled = true;

            // Create a request to convert identifiers. 
            AlternateId objAltID = new AlternateId();
            objAltID.Format = IdFormat.HexEntryId;
            objAltID.Mailbox = "vinod@dev.local";
            objAltID.IsArchive = true;
            objAltID.UniqueId = sID;

            //Convert  PR_ENTRYID identifier format to an EWS identifier. 
            AlternateIdBase objAltIDBase = service.ConvertId(objAltID, IdFormat.EwsId);
            AlternateId objAltIDResp = (AlternateId)objAltIDBase;

            Console.WriteLine(objAltIDResp.UniqueId);
            //sEwsId1 = GetConvertedEntryID(service, "AAMkAGFkZTM1MjY3LWZiYzAtNDA1ZC04NWI3LTA1ZWRlYzE2NjVjZABGAAAAAAAehyvl2c+VRaNBUFlASUlpBwA5Thqx2ogYS5z4GmODBiBuAAF/PIMQAAA5Thqx2ogYS5z4GmODBiBuAAGi4SjGAAA=",
              //                          smtpAddress);
        }

        public static void BindFolder(string sEwsId)
        {
            ExchangeService service;
            service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);


            service.Credentials = new WebCredentials("ewsuser@dev.local", "!manage32");
            service.TraceListener = new TraceListener();
            service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

            string smtpAddress = "vinod@dev.local";
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

            string exchangeUrl;
            exchangeUrl = "https://";
            exchangeUrl += "xchange.dev.local";
            exchangeUrl += "/EWS/Exchange.asmx";


            service.Url = new Uri(exchangeUrl);


            Console.WriteLine("AutodiscoverURL: " + service.Url);

            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;

            service.TraceEnabled = true;

            Folder fld;
            
            Console.WriteLine("Id: " + sEwsId);
            FolderId id = new FolderId(sEwsId);

            fld = Folder.Bind(service, id);
            Console.WriteLine("");
            Console.WriteLine("Folder Name: " + fld.DisplayName);

            ////////////////////

            if (fld.DisplayName.Length > 0)
            {
                var view = new ItemView(100) { PropertySet = new PropertySet { EmailMessageSchema.Id, ItemSchema.Subject } };
                view.Traversal = ItemTraversal.Shallow;
                FindItemsResults<Item> findResults = service.FindItems(fld.Id, view);
                Console.Write("Total number of emails " + findResults.TotalCount);

                foreach (Item myItem in findResults.Items)
                {
                    if (myItem is EmailMessage)
                    {
                        Console.WriteLine((myItem as EmailMessage).Subject);
                    }
                }
            }

        }


        public static void FindFolder(int iType)
        {

            ///////////////////////////Find folders in Archive mailbox -start/////////////////
             ExchangeService service;
            service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);


            service.Credentials = new WebCredentials("ewsuser@dev.local", "!manage32");
            service.TraceListener = new TraceListener();
            service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

            string smtpAddress = "vinod@dev.local";
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

            string exchangeUrl;
            exchangeUrl = "https://";
            exchangeUrl += "xchange.dev.local";
            exchangeUrl += "/EWS/Exchange.asmx";


            service.Url = new Uri(exchangeUrl);


            Console.WriteLine("AutodiscoverURL: " + service.Url);

            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;

            service.TraceEnabled = true;

            FolderView view = new FolderView(100);
            view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
            view.PropertySet.Add(FolderSchema.DisplayName);
            view.Traversal = FolderTraversal.Deep;
            FindFoldersResults findFolderResults ;//= service.FindFolders(WellKnownFolderName.ArchiveRoot, view);
            if (iType > 0)
            {
                Console.WriteLine("Archive Root");
                findFolderResults = service.FindFolders(WellKnownFolderName.ArchiveRoot, view);
            }
            else
            {
                Console.WriteLine("mailbox root");
                findFolderResults = service.FindFolders(WellKnownFolderName.Root, view);
            }
            //find specific folder
            foreach (Folder f in findFolderResults)
            {
                //show folderId of the folder "test"
                if (f.DisplayName == "Archive Folder")
                {
                    Console.WriteLine("Found Archive Folder ");
                    Console.WriteLine(f.Id);
                }
                //else
                //    Console.WriteLine(f.DisplayName);
            }
            ///////////////////////////Find folders in Archive mailbox -End/////////////////
            

        }
        

        public static void FindItemEmail(int iType)
        {

            ///////////////////////////Find folders in Archive mailbox -start/////////////////
            // ExchangeService service;
            //service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);


            ////service.Credentials = new WebCredentials("admin@imanage.microsoftonline.com", "!wov2014");
            //service.Credentials = new WebCredentials("ewsuser@dev.local", "!manage32");
            //service.TraceListener = new TraceListener();
            //service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

            ////string smtpAddress = "kenneth.lay@imanage.microsoftonline.com";
            //string smtpAddress = "vinod@dev.local";
            //service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

            //string exchangeUrl;
            //exchangeUrl = "https://";
            ////exchangeUrl += "ch1prd0410.outlook.com";
            //exchangeUrl += "xchange.dev.local";
            //exchangeUrl += "/EWS/Exchange.asmx";


            //service.Url = new Uri(exchangeUrl);


            //Console.WriteLine("AutodiscoverURL: " + service.Url);

            //ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;

            //service.TraceEnabled = true;

            //FolderView view = new FolderView(100);
            //view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
            //view.PropertySet.Add(FolderSchema.DisplayName);
            //view.Traversal = FolderTraversal.Deep;
            //FindFoldersResults findFolderResults ;//= service.FindFolders(WellKnownFolderName.ArchiveRoot, view);
            //if (iType > 0)
            //{
            //    Console.WriteLine("Archive Root");
            //    findFolderResults = service.FindFolders(WellKnownFolderName.ArchiveRoot, view);
            //}
            //else
            //{
            //    Console.WriteLine("mailbox root");
            //    findFolderResults = service.FindFolders(WellKnownFolderName.Root, view);
            //}
            ////find specific folder
            //foreach (Folder f in findFolderResults)
            //{
            //    //show folderId of the folder "test"
            //    if (f.DisplayName == "ArchFold1")
            //    {
            //        Console.WriteLine("Found Archive Fold ");
            //        Console.WriteLine(f.Id);
            //    }
            //}
            ///////////////////////////Find folders in Archive mailbox -End/////////////////
/////////////////// Find items in Archive mailbox -start/////////////////
            ExchangeService service;
            service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);


            service.Credentials = new WebCredentials("ewsuser@dev.local", "!manage32");
            service.TraceListener = new TraceListener();
            service.TraceFlags = TraceFlags.All;// TraceFlags.EwsRequest | TraceFlags.EwsResponse;

            
            string smtpAddress = "vinod@dev.local";
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);

            string exchangeUrl;
            exchangeUrl = "https://";
            exchangeUrl += "xchange.dev.local";
            exchangeUrl += "/EWS/Exchange.asmx";


            service.Url = new Uri(exchangeUrl);


            Console.WriteLine("AutodiscoverURL: " + service.Url);

            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;

            service.TraceEnabled = true;

            
             FolderView folderView = new FolderView(1000);
            folderView.Traversal = FolderTraversal.Shallow;
            
            //FolderId rootFolderId = new FolderId(WellKnownFolderName.Inbox);
            //SearchFilter.SearchFilterCollection searchFilterCollection = new SearchFilter.SearchFilterCollection();            
            //SearchFilter searchFilter1 = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "2");
            ////SearchFilter searchFilter1 = new SearchFilter.IsEqualTo(FolderSchema.Id, "AAMkAGFkZTM1MjY3LWZiYzAtNDA1ZC04NWI3LTA1ZWRlYzE2NjVjZAAuAAAAAAAehyvl2c+VRaNBUFlASUlpAQA5Thqx2ogYS5z4GmODBiBuAABACwAgAAA=");
            //searchFilterCollection.Add(searchFilter1);



            //FindFoldersResults findFoldersResults = service.FindFolders(rootFolderId, searchFilterCollection, folderView);
            //if (findFoldersResults.Folders.Count > 0)
            //{
            //    Folder fold = findFoldersResults.Folders[0];
            //    Console.WriteLine("Folder:\t" + fold.DisplayName);
            //    Console.WriteLine("Folder:\t" + fold.Id);
            //}
            ////////////////////////
            ExtendedPropertyDefinition prFolderType = new ExtendedPropertyDefinition(13825, MapiPropertyType.Integer);
            SearchFilter.SearchFilterCollection filterAllItemsFolder = new SearchFilter.SearchFilterCollection(LogicalOperator.And);
            filterAllItemsFolder.Add(new SearchFilter.IsEqualTo(prFolderType, "2"));
            filterAllItemsFolder.Add(new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "AllItems"));
            FolderView viewAllItemsFolder = new FolderView(1000);
            viewAllItemsFolder.Traversal = FolderTraversal.Shallow;

            FindFoldersResults findAllItemsFolder;
            if (iType > 0)
            {
                Console.WriteLine("Archive Root");
                findAllItemsFolder = service.FindFolders(WellKnownFolderName.ArchiveRoot, filterAllItemsFolder, viewAllItemsFolder);
            }
            else
            {
                Console.WriteLine("mailbox root");
                findAllItemsFolder = service.FindFolders(WellKnownFolderName.Root, filterAllItemsFolder, viewAllItemsFolder);
            }

            if (findAllItemsFolder.Folders.Count > 0)
            {
                Folder allItemsFolder = findAllItemsFolder.Folders[0];
                Console.WriteLine("Folder:\t" + allItemsFolder.DisplayName);

                String searchstring = "RE: new bits";// "test march 20 - 001";
                SearchFilter.IsEqualTo filter = new SearchFilter.IsEqualTo(EmailMessageSchema.Subject, searchstring);

                var view = new ItemView(100) { PropertySet = new PropertySet { EmailMessageSchema.Id, ItemSchema.Subject } };
                view.Traversal = ItemTraversal.Shallow;
                FindItemsResults<Item> findResults = service.FindItems(allItemsFolder.Id, filter, view);
                Console.Write("IsEqualTo: Total email count with the specified search string in the subject: " + findResults.TotalCount);

                foreach (Item myItem in findResults.Items)
                {
                    if (myItem is EmailMessage)
                    {
                        Console.WriteLine((myItem as EmailMessage).Subject);
                    }
                }
            }
///////////////////////Find items in Archive mailbox - end////////////////////

        }

       

        // BindFolder
        
        private static String GetConvertedEntryID(ExchangeService esb, String sID, String strSMTPAdd)
        {
            // Create a request to convert identifiers. 
            AlternateId objAltID = new AlternateId();
            objAltID.Format = IdFormat.EwsId;
            objAltID.Mailbox = strSMTPAdd;
            objAltID.UniqueId = sID;

            //Convert  PR_ENTRYID identifier format to an EWS identifier. 
            AlternateIdBase objAltIDBase = esb.ConvertId(objAltID, IdFormat.HexEntryId);
            AlternateId objAltIDResp = (AlternateId)objAltIDBase;
            return objAltIDResp.UniqueId;
        }

        
        private static String GetConvertedEWSID(ExchangeService esb, String sID, String strSMTPAdd, bool bIsArchive)
        {
            // Create a request to convert identifiers. 
            AlternateId objAltID = new AlternateId();
            objAltID.Format = IdFormat.HexEntryId;
            objAltID.Mailbox = strSMTPAdd;
            objAltID.IsArchive = bIsArchive;
            objAltID.UniqueId = sID;

            //Convert  PR_ENTRYID identifier format to an EWS identifier. 
            AlternateIdBase objAltIDBase = esb.ConvertId(objAltID, IdFormat.EwsId);
            AlternateId objAltIDResp = (AlternateId)objAltIDBase;
            return objAltIDResp.UniqueId;
        }

        

         
        static void Main(string[] args)
        {
            string sCmd = "6";// args[0];
            if (sCmd == "1")
            {
                //if (args.Length > 0)
                    FindItemEmail(1);
                //else
                //    FindItemEmail(0);
                return;
            }

            if (sCmd == "2")
            {
                //if (args.Length > 0)
                    FindFolder(1);
                //else
                //    FindFolder(0);
                return;
            }

            if (sCmd == "3")
            {
                if (args.Length > 1)
                {
                    string sEwsId = args[1];
                    BindFolder(sEwsId);
                }
            }

            if (sCmd == "4")
            {
                if (args.Length > 1)
                {
                    string sEntryId = args[1];
                    GetEWSIdFromEntryId(sEntryId);
                }
            }

            if (sCmd == "5")
            {
                if (args.Length > 1)
                {
                    string sEwsId = args[1];
                    GetEntryIdFromEWSId(sEwsId);
                    
                }
            }

            if (sCmd == "6")
            {
                if (args.Length > 1)
                {
                    string sEwsId = args[1]; //"AAMkAGFkZTM1MjY3LWZiYzAtNDA1ZC04NWI3LTA1ZWRlYzE2NjVjZAAuAAAAAAAehyvl2c+VRaNBUFlASUlpAQA5Thqx2ogYS5z4GmODBiBuAAFZ+vdlAAA=";//args[1];
                    GetItemsFromFolders(args);

                }
            }

            ////FindUnMarkedItem();
            //if (args.Length > 0)
            //{
            //    for (int i = 0; i < args.Length; i++ )
            //    {
            //        Console.WriteLine(args[i]);
            //    }

               
            //    Console.WriteLine("Done");
            //}
            //else
            //{
            //    Console.WriteLine("Invalid Command");
            //    Console.WriteLine("Available commands 1) Autodiscover 2) TestConnection 3) BindFolder 4) PropCleanup");
            //    Console.WriteLine("for more info on a particular command type EWSTestApp CommandName");
            //}
        }
               
    }
}
            

