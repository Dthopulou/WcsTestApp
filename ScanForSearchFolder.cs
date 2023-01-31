using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
namespace EWSTestApp
{
    class ScanForSearchFolder
    {
        public void ScanSearchFolder(string[] args)
        {

            StreamWriter Log;
            if (!File.Exists("EWSTestAppLog1.txt"))
            {
                Log = new StreamWriter("EWSTestAppLog1.txt", true);
            }
            else
            {
                Log = File.AppendText("EWSTestAppLog1.txt");

            }

            Log.AutoFlush = true;

            StreamWriter MsgId;
            if (!File.Exists("MessageId.txt"))
            {
                MsgId = new StreamWriter("MessageId.txt", true);
            }
            else
            {
                MsgId = File.AppendText("MessageId.txt");

            }

            MsgId.AutoFlush = true;

            if (args.Length < 8)
            {
                Console.WriteLine("Syntax: <Command> <ImpersonatorSMTP> <password> <exchange server name> <ExchangeVersion> <SearchFolderParent> <SearchFolder> <sReportmode>");
                Console.WriteLine("Example: SCAN_SEARCH_FOLDER ImpersonatorSMTPAddress@dev.local password exchangeServer Exchange2010 1 1 False/True");

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
                string sReportMode = args[7].ToUpper();
                string RenameFolderOrDelete = args[8];
                System.IO.StreamReader file = new System.IO.StreamReader("Users.txt");
                string line;
                //long iTotalEmailCount = 0;
                while ((line = file.ReadLine()) != null)
                {
                    try
                    {
                        if (line.Length > 0)
                        {
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
                            FolderView folderView = new FolderView(50);
                            folderView.PropertySet = new PropertySet(FolderSchema.DisplayName);
                            //int itemcount = 0;
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
                            //else if (sSearchFoldParent == "3")
                            //    wellknownFoldName = WellKnownFolderName.SearchFolders;
                            else
                                wellknownFoldName = WellKnownFolderName.Root;
                            //if (sReportMode == "TRUE")
                            //{
                            //    Log.WriteLine("Generic Folder will be deleted if it is empty and renamed/Delete if it has item");
                            //}
                            //else
                            {
                                //FindItemsResults<Item> findItemResults;
                                //ItemView view = new ItemView(10);
                                //ExtendedPropertyDefinition PR_SENT_REPRESENTING = new ExtendedPropertyDefinition(0x5D02, MapiPropertyType.String);
                                //ExtendedPropertyDefinition PR_SENDR_NAME = new ExtendedPropertyDefinition(0x0C1A, MapiPropertyType.String);
                                
                               // view.PropertySet = new PropertySet(PR_SENT_REPRESENTING,PR_SENDR_NAME);
                               //// view.Traversal = ItemTraversal.Shallow;
                               // //SearchFilter.SearchFilterCollection dateCriAndCollection =
                               // //                                              new SearchFilter.SearchFilterCollection(LogicalOperator.And);
                               // //ExtendedPropertyDefinition PR_SENT_REPRESENTING = new ExtendedPropertyDefinition(0x0042, MapiPropertyType.String);
                               // //ExtendedPropertyDefinition PR_SENDR_NAME = new ExtendedPropertyDefinition(0x0C1A, MapiPropertyType.String);
                               // //dateCriAndCollection.Add(new SearchFilter.ContainsSubstring(""));
                               //// dateCriAndCollection.Add(new SearchFilter.Exists(PR_SENDR_NAME));
                               // SearchFilter.ContainsSubstring searchCriteria =new SearchFilter.ContainsSubstring(ItemSchema.Subject, "Test for final delegate");
                               // findItemResults = service.FindItems(WellKnownFolderName.Inbox, searchCriteria, view);
                               // Console.WriteLine("Processing User: {0}", smtpAddress);
                               // Log.WriteLine("Processing User: {0}", smtpAddress);
                               // FindFoldersResults findFoldResults = service.FindFolders(wellknownFoldName, searchFoldFilter, folderView);

                               // Console.WriteLine("Processing User: {0}", smtpAddress);
                               // Log.WriteLine("Processing User: {0}", smtpAddress);
                               // //FindFoldersResults findFoldResults = service.FindFolders(wellknownFoldName, searchFoldFilter, folderView);

                               // if (findFoldResults.Count() == 0)
                               // {
                               //     Console.WriteLine("{0} has no folder with the name {1}", smtpAddress, sFolderName);
                               //     Log.WriteLine("{0} has no folder with the name {1}", smtpAddress, sFolderName);
                               // }

                              // ----------Test for PR_SENT_REPRESENTING item Folder--------


                               // FindItemsResults<Item> findResults;
                                FindFoldersResults findFoldResults = service.FindFolders(wellknownFoldName, folderView);
                                foreach (Folder folder in findFoldResults.Folders)
                                {
                                   // Console.WriteLine("User: {0} Processed Successfully", folder.DisplayName);
                                    if (folder.DisplayName.Equals("Sent Items"))
                                    {
                                    FindItemsResults<Item> findResults;
                                    ItemView view = new ItemView(50, 0, OffsetBasePoint.Beginning);                                   
                                    view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
                                    view.Traversal = ItemTraversal.Shallow;
                                    

                                    SearchFilter.SearchFilterCollection searchOrFilterCollection = new SearchFilter.SearchFilterCollection(LogicalOperator.And);

                                  searchOrFilterCollection.Add(new SearchFilter.IsLessThan(ItemSchema.DateTimeReceived, DateTime.Today.AddDays(+30)));
                                   //searchOrFilterCollection.Add(new SearchFilter.IsLessThan(ItemSchema.DateTimeSent, DateTime.Today.AddDays(-30)));
                                   searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.WorkSite.Ems.Filed"));
                                   searchOrFilterCollection.Add(new SearchFilter.ContainsSubstring(ItemSchema.Subject, "["));
                                   

                                    findResults = service.FindItems(folder.Id, searchOrFilterCollection, view);
                                        foreach (Item myItem in findResults.Items)
                                        {
                                            //EmailMessage em = EmailMessage.Bind(service, myItem.Id, new PropertySet(EmailMessageSchema.InternetMessageId));
                                            //if (em is EmailMessage)
                                            //{
                                            //    string sMsgid = em.InternetMessageId;
                                            //    MsgId.WriteLine("MessageId: {0}", sMsgid);
                                            //    Console.WriteLine(em.InternetMessageId);
                                            //}
                                            EmailMessage message = EmailMessage.Bind(service, myItem.Id.UniqueId, new PropertySet(ItemSchema.Attachments));
                                            foreach (Attachment attachment in message.Attachments)
                                            {
                                                if (attachment is ItemAttachment)
                                                {
                                                    ItemAttachment itemAttachment = attachment as ItemAttachment;
                                                    itemAttachment.Load(ItemSchema.MimeContent);
                                                    string fileName = "C:\\Temp\\" + itemAttachment.Item.Subject + ".eml";
                                                    // Write the bytes of the attachment into a file.
                                                    File.WriteAllBytes(fileName, itemAttachment.Item.MimeContent.Content);
                                                    Console.WriteLine("Email attachment name: " + itemAttachment.Item.Subject + ".eml");
                                                    EmailMessage em = EmailMessage.Bind(service, myItem.Id, new PropertySet(EmailMessageSchema.InternetMessageId));
                                                    if (em is EmailMessage)
                                                    {
                                                        string sMsgid = em.InternetMessageId;
                                                        MsgId.WriteLine("MessageId: {0}", sMsgid);
                                                        Console.WriteLine(em.InternetMessageId);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                
                                }
                                
                              


                                //foreach (Folder folder in findFoldResults.Folders)
                                //{
                                //    if (folder is SearchFolder)
                                //    {
                                //        Log.WriteLine("{0} is a search folder", folder.DisplayName);
                                //        FindItemsResults<Item> EmailItems;
                                //        ItemView view = new ItemView(5);
                                //        view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
                                //        view.Traversal = ItemTraversal.Shallow;
                                //        EmailItems = service.FindItems(folder.Id, view);

                                //        if (EmailItems.Count() == 0)
                                //        {
                                //            Folder folder2 = Folder.Bind(service, folder.Id);
                                //            if (sReportMode != "TRUE")
                                //                folder.Delete(DeleteMode.HardDelete);
                                //            Console.WriteLine("Folder : {0} is not a WCSE Search folder and it has been deleted for User: {1}", folder2.DisplayName, smtpAddress);
                                //            Log.WriteLine("Folder : {0} is not a WCSE Search folder and it has been deleted for User: {1}", folder2.DisplayName, smtpAddress);
                                //        }
                                //    }
                                //    else
                                //    {
                                //        FindItemsResults<Item> EmailItems;
                                //        ItemView view = new ItemView(5);
                                //        view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
                                //        view.Traversal = ItemTraversal.Shallow;
                                //        EmailItems = service.FindItems(folder.Id, view);

                                //        //if (sReportMode == "TRUE")
                                //        //{
                                //        //    Log.WriteLine("Folder will be deleted if it is empty and renamed if it has item");
                                //        //}
                                //        //else
                                //        {
                                //            if (EmailItems.Count() == 0)
                                //            {
                                //                Folder folder2 = Folder.Bind(service, folder.Id);
                                //                if (sReportMode != "TRUE")
                                //                    folder2.Delete(DeleteMode.HardDelete);
                                //                Console.WriteLine("Folder : {0} has  {1} emails", folder2.DisplayName, EmailItems.Count());
                                //                Console.WriteLine("Folder : {0} is not a WCSE Search folder and it has been deleted for User: {1}", folder2.DisplayName, smtpAddress);
                                //                Log.WriteLine("Folder : {0} is not a WCSE Search folder and it has been deleted for User: {1}", folder2.DisplayName, smtpAddress);
                                //            }
                                //            else
                                //            {
                                //                if (RenameFolderOrDelete.ToUpper() == "TRUE")
                                //                {
                                //                    Folder folder2 = Folder.Bind(service, folder.Id);
                                //                    if (sReportMode != "TRUE")
                                //                        folder2.Delete(DeleteMode.HardDelete);
                   
                                //                    Console.WriteLine("Folder : {0} has  {1} emails", folder2.DisplayName, EmailItems.Count());
                                //                    Console.WriteLine("Folder : {0} is not a WCSE Search folder and it has been deleted for User: {1}", folder2.DisplayName, smtpAddress);
                                //                    Log.WriteLine("Folder : {0} is not a WCSE Search folder and it has been deleted for User: {1}", folder2.DisplayName, smtpAddress);

                                //                }
                                //                else
                                //                {
                                //                    Console.WriteLine("Folder : {0} has  {1} emails", folder.DisplayName, EmailItems.Count());
                                //                    Log.WriteLine("Folder : {0} has  {1} emails", folder.DisplayName, EmailItems.Count());
                                //                    //update folder name if generic folder is same as wcse search folder
                                //                    // As a best practice, only include the ID value in the PropertySet.
                                //                    PropertySet propertySet = new PropertySet(BasePropertySet.IdOnly);

                                //                    // Bind to an existing folder and get the FolderId.
                                //                    // This method call results in a GetFolder call to EWS.
                                //                    Folder folder1 = Folder.Bind(service, folder.Id, propertySet);

                                //                    // Update the display name of the folder.
                                //                    folder1.DisplayName = RenameFolderOrDelete;


                                //                    // Save the updates.
                                //                    // This method call results in an UpdateFolder call to EWS.
                                //                    if (sReportMode != "TRUE")
                                //                        folder1.Update();
                                //                    Console.WriteLine("Generic SearchFodler:{0} Renamed to {1} for User:{2}", sFolderName, folder1.DisplayName, smtpAddress);
                                //                    Log.WriteLine("Generic SearchFodler:{0} Renamed to {1} for User:{2}",sFolderName, folder1.DisplayName, smtpAddress);
                                //                    //Log.WriteLine("Generic SearchFolder renamed to {1}", folder1.DisplayName);
                                //                }

                                //            }

                                //        }



                                //    }

                                //}
                                Console.WriteLine("User: {0} Processed Successfully", smtpAddress);
                                Log.WriteLine("User: {0} Processed Successfully", smtpAddress);
                            }

                        }

                    }
                    catch (Exception ex)
                    {

                        DateTime dt = DateTime.Now;
                        Log.WriteLine("Folder: {0} - {1} ", dt, ex.Message);
                        Console.WriteLine("{0}",ex.Message);
                    }
                }
                Log.Close();

            } while (false);
        }
    }
}
