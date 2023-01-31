using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
namespace EWSTestApp
{
    class ScanSentItemFolderForSendAndFile
    {
        public void ScanSentItemFolder(string[] args)
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

            if (args.Length < 7)
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
                string sFirmId = args[5];
                string sTimeToRetrive = args[6];
                string sReportMode = args[7].ToUpper();
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
                            //FolderView folderView = new FolderView(50);
                            //folderView.PropertySet = new PropertySet(FolderSchema.DisplayName);
                            ////int itemcount = 0;
                            //SearchFilter searchFoldFilter = null;
                            //string sFolderName = "";


                            //if (sSearchFold == "1")
                            //{
                            //    sFolderName = "WCSE_FolderMappings";
                            //    searchFoldFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, sFolderName);
                            //}
                            //else if (sSearchFold == "2")
                            //{
                            //    sFolderName = "WCSE_SFMailboxSync";
                            //    searchFoldFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, sFolderName);
                            //}

                            //WellKnownFolderName wellknownFoldName;
                            //if (sSearchFoldParent == "1")
                            //    wellknownFoldName = WellKnownFolderName.Root;
                            //else if (sSearchFoldParent == "2")
                            //    wellknownFoldName = WellKnownFolderName.MsgFolderRoot;
                            //else
                            //    wellknownFoldName = WellKnownFolderName.Root;
                            if (sReportMode == "TRUE")
                            {
                                Log.WriteLine("All the Filling informations(filingDocumentId, filingStatus, filingLocation, filingStatusCode, filingfolder, filingDate, filingFolderId) will be deleted");
                            }
                            else
                            {
                                //FindFoldersResults findFoldResults = service.FindFolders(wellknownFoldName, folderView);
                                //foreach (Folder folder in findFoldResults.Folders)
                               // {
                                    //if (folder.DisplayName.Equals("Sent Items"))
                                    //{
                                string ExProp = "";
                                string sDocnum = "";
                                int lDocnum;
                                StreamWriter OuputCSV = null;
                                if (File.Exists("DocNum.txt"))
                                {
                                    File.Delete("DocNum.txt");
                                    OuputCSV = new StreamWriter("DocNum.txt", true);
                                    OuputCSV.AutoFlush = true;
                                }
                                else
                                {
                                    OuputCSV = new StreamWriter("DocNum.txt", true);
                                    OuputCSV.AutoFlush = true;
                                }
                                        FindItemsResults<Item> findResults;
                                        ExtendedPropertyDefinition filingDocumentId = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                       "FilingDocumentID", MapiPropertyType.String);
                                        ExtendedPropertyDefinition filingStatus = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                              "FilingStatus", MapiPropertyType.String);
                                        ExtendedPropertyDefinition filingLocation = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                              "FilingLocation", MapiPropertyType.String);
                                        ExtendedPropertyDefinition filingStatusCode = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                              "FilingStatusCode", MapiPropertyType.Integer);
                                        ExtendedPropertyDefinition filingfolder = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                              "FilingFolder", MapiPropertyType.String);
                                        ExtendedPropertyDefinition filingDate = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                              "FilingDate", MapiPropertyType.SystemTime);
                                        ExtendedPropertyDefinition filingFolderId = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings,
                                                                                                      "FilingFolderID", MapiPropertyType.String);
                                        ExtendedPropertyDefinition PR_Search_Key = new ExtendedPropertyDefinition(0x300B, MapiPropertyType.Binary);
                                       
     
                                        ItemView view = new ItemView(5, 0, OffsetBasePoint.Beginning);
                                        view.PropertySet = new PropertySet(BasePropertySet.IdOnly, filingDocumentId, filingStatus, filingLocation, filingStatusCode, filingfolder, filingDate, filingFolderId);
                                        //PropertySet prop = BasePropertySet.IdOnly;
                                        //prop.Add(PR_Search_Key);
                                        //ItemView ivItemView = new ItemView(10);
                                        //view.PropertySet = prop;
                      

                                        SearchFilter.SearchFilterCollection searchOrFilterCollection = new SearchFilter.SearchFilterCollection(LogicalOperator.And);
                                        Double dTimeToRetrive = Convert.ToDouble(sTimeToRetrive);
                                        searchOrFilterCollection.Add(new SearchFilter.IsGreaterThan(ItemSchema.DateTimeReceived, DateTime.Today.AddDays(-dTimeToRetrive)));
                                       // searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.DateTimeReceived, DateTime.Today));
                                        searchOrFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note.WorkSite.Ems.Filed"));
                                        searchOrFilterCollection.Add(new SearchFilter.ContainsSubstring(ItemSchema.Subject, sFirmId));

                                      
                                        findResults = service.FindItems(WellKnownFolderName.SentItems, searchOrFilterCollection, view);
                                        //foreach (Item item in findResults)
                                        //{
                                        //    Byte[] PropVal;
                                        //    String HexSearchKey;
                                        //    if (item.TryGetProperty(PR_Search_Key, out PropVal))
                                        //    {
                                        //        HexSearchKey = BitConverter.ToString(PropVal).Replace("-", "");
                                        //    }

                                        //}
                                        Byte[] PropVal;
                                        String HexSearchKey;
                                        Item message1 = null;
                                
                                        foreach (Item myItem in findResults.Items)
                                        {
                                            //if (myItem.ExtendedProperties.Count > 0)
                                            //{

                                            foreach (ExtendedProperty extendedProperty in myItem.ExtendedProperties)
                                            {
                                                ExProp = extendedProperty.PropertyDefinition.Name.ToString();
                                                if (ExProp.Equals("FilingDocumentID"))
                                                // if (extendedProperty.PropertyDefinition.Name.Equals("FilingDocumentID"))
                                                {

                                                    string sFilingDocumentId= extendedProperty.Value.ToString();
                                                    int start = sFilingDocumentId.IndexOf("document") + 9;
                                                    int end = sFilingDocumentId.LastIndexOf(":") - 2;
                                                    sDocnum = sFilingDocumentId.Substring(start, end - start);
                                                    lDocnum = Int32.Parse(sDocnum);
                                                    OuputCSV.WriteLine("{0}", lDocnum);
                                                    Console.WriteLine("docnum {0}:", lDocnum);
                                                }

                                            }
                                            //}
                                            if (myItem.ExtendedProperties.Count > 0)
                                            {
                                                // Display the extended name and value of the extended property.
                                                foreach (ExtendedProperty extendedProperty in myItem.ExtendedProperties)
                                                {
                                                    if ((extendedProperty.PropertyDefinition.Name == filingDocumentId.Name && extendedProperty.PropertyDefinition.PropertySetId == filingDocumentId.PropertySetId))
                                                    {
                                                        message1 = myItem;
                                                        message1.RemoveExtendedProperty(filingDocumentId);
                                                        message1.Update(ConflictResolutionMode.AlwaysOverwrite);
                                                        break;
                                                    }
                                                }

                                                foreach (ExtendedProperty extendedProperty in myItem.ExtendedProperties)
                                                {
                                                    if ((extendedProperty.PropertyDefinition.Name == filingStatus.Name && extendedProperty.PropertyDefinition.PropertySetId == filingStatus.PropertySetId))
                                                    {
                                                        message1 = myItem;
                                                        message1.RemoveExtendedProperty(filingStatus);
                                                        message1.Update(ConflictResolutionMode.AlwaysOverwrite);
                                                        break;
                                                    }
                                                }

                                                foreach (ExtendedProperty extendedProperty in myItem.ExtendedProperties)
                                                {
                                                    if ((extendedProperty.PropertyDefinition.Name == filingLocation.Name && extendedProperty.PropertyDefinition.PropertySetId == filingLocation.PropertySetId))
                                                    {
                                                        message1 = myItem;
                                                        message1.RemoveExtendedProperty(filingLocation);
                                                        message1.Update(ConflictResolutionMode.AlwaysOverwrite);
                                                        break;
                                                    }
                                                }

                                                foreach (ExtendedProperty extendedProperty in myItem.ExtendedProperties)
                                                {
                                                    if ((extendedProperty.PropertyDefinition.Name == filingStatusCode.Name && extendedProperty.PropertyDefinition.PropertySetId == filingStatusCode.PropertySetId))
                                                    {
                                                        message1 = myItem;
                                                        message1.RemoveExtendedProperty(filingStatusCode);
                                                        message1.Update(ConflictResolutionMode.AlwaysOverwrite);
                                                        break;
                                                    }
                                                }

                                                foreach (ExtendedProperty extendedProperty in myItem.ExtendedProperties)
                                                {
                                                    if ((extendedProperty.PropertyDefinition.Name == filingfolder.Name && extendedProperty.PropertyDefinition.PropertySetId == filingfolder.PropertySetId))
                                                    {
                                                        message1 = myItem;
                                                        message1.RemoveExtendedProperty(filingfolder);
                                                        message1.Update(ConflictResolutionMode.AlwaysOverwrite);
                                                        break;
                                                    }
                                                }

                                                foreach (ExtendedProperty extendedProperty in myItem.ExtendedProperties)
                                                {
                                                    if ((extendedProperty.PropertyDefinition.Name == filingDate.Name && extendedProperty.PropertyDefinition.PropertySetId == filingDate.PropertySetId))
                                                    {
                                                        message1 = myItem;
                                                        message1.RemoveExtendedProperty(filingDate);
                                                        message1.Update(ConflictResolutionMode.AlwaysOverwrite);
                                                        break;
                                                    }
                                                }

                                                foreach (ExtendedProperty extendedProperty in myItem.ExtendedProperties)
                                                {
                                                    if ((extendedProperty.PropertyDefinition.Name == filingFolderId.Name && extendedProperty.PropertyDefinition.PropertySetId == filingFolderId.PropertySetId))
                                                    {
                                                        message1 = myItem;
                                                        message1.RemoveExtendedProperty(filingFolderId);
                                                        message1.Update(ConflictResolutionMode.AlwaysOverwrite);
                                                        break;
                                                    }
                                                }

                                               
                                            }
                                            PropertySet props = new PropertySet(BasePropertySet.IdOnly, EmailMessageSchema.MimeContent, ItemSchema.Subject, ItemSchema.Attachments,
                                                ItemSchema.HasAttachments, ItemSchema.ItemClass, ItemSchema.Size, ItemSchema.DateTimeReceived, ItemSchema.DateTimeSent, filingDocumentId, filingStatus, filingLocation, filingStatusCode, filingfolder, filingDate, filingFolderId, PR_Search_Key);
                                            
                                            Item item = EmailMessage.Bind(service, myItem.Id, props);
                                           
                                            
                                           if (item.Attachments.Count > 0)
                                            {
                                                if (myItem.ExtendedProperties.Count > 0)
                                                {

                                                    //foreach (ExtendedProperty extendedProperty in myItem.ExtendedProperties)
                                                    //{
                                                    //    ExProp = extendedProperty.Value.ToString();
                                                    //    if (ExProp.Equals("FilingDocumentID"))
                                                    //    // if (extendedProperty.PropertyDefinition.Name.Equals("FilingDocumentID"))
                                                    //    {
                                                    //        // FillingDocId = extendedProperty.Value.ToString();
                                                    //        int start = ExProp.IndexOf("document") + 9;
                                                    //        int end = ExProp.LastIndexOf(":") - 2;
                                                    //        sDocnum = ExProp.Substring(start, end - start);
                                                    //        lDocnum = Int32.Parse(sDocnum);
                                                    //        OuputCSV.WriteLine("{0}", lDocnum);
                                                    //        Console.WriteLine("docnum {0}:", lDocnum);
                                                    //    }

                                                    //}
                                                }

                                                if (item.TryGetProperty(PR_Search_Key, out PropVal))
                                                   {
                                                       HexSearchKey = BitConverter.ToString(PropVal).Replace("-", "");

                                                       String CurrentDir = Directory.GetCurrentDirectory();
                                                       //string emlFileName = @"c:\export\" + HexSearchKey + ".eml";

                                                       String EmailsFolder = Path.Combine(CurrentDir, "export");

                                                       if (!Directory.Exists(EmailsFolder))
                                                       {
                                                           Directory.CreateDirectory(EmailsFolder);
                                                           if (!Directory.Exists(EmailsFolder))
                                                           {
                                                               Console.WriteLine("Failed to create E-Mails folder under " + CurrentDir);
                                                               break;
                                                           }
                                                       }
                                                       string emlFileName = String.Format(@"{0}\export\{1}.eml", CurrentDir, HexSearchKey);
                                                       //string emlFileName = @"CurrentDir\" + HexSearchKey + ".eml";
                                                       using (FileStream fs = new FileStream(emlFileName, FileMode.Create, FileAccess.Write))
                                                       {
                                                           fs.Write(item.MimeContent.Content, 0, item.MimeContent.Content.Length);
                                                       }
                                                   }

                                            }
                    
                                        }
                                   // }

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
                        Console.WriteLine("{0}", ex.Message);
                    }
                }
                Log.Close();

            } while (false);
        }
    }
}
