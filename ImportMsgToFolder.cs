using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using System.Runtime.InteropServices;
using Redemption;

namespace EWSTestApp
{
    class ImportMsgToFolder
    {
        [DllImport("kernel32.dll")]
        static extern IntPtr LoadLibrary(string dllName);

        [DllImport("kernel32.dll")]
        static extern IntPtr GetProcAddress(IntPtr hModule, string procName);

        delegate int MessageBoxDelegate(IntPtr hwnd,
            [MarshalAs(UnmanagedType.LPWStr)]string text,
            [MarshalAs(UnmanagedType.LPWStr)]string caption,
            int type);
        public void ImportMsgToSpecifiedFolder(string[] args)
        {

            StreamWriter Log;
            IntPtr mRedemption;
            if (!File.Exists("EWSTestAppLog1.txt"))
            {
                Log = new StreamWriter("EWSTestAppLog1.txt", true);
            }
            else
            {
                Log = File.AppendText("EWSTestAppLog1.txt");

            }

            Log.AutoFlush = true;

            //StreamWriter MsgId;
            //if (!File.Exists("MessageId.txt"))
            //{
            //    MsgId = new StreamWriter("MessageId.txt", true);
            //}
            //else
            //{
            //    MsgId = File.AppendText("MessageId.txt");

            //}

            //MsgId.AutoFlush = true;

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
                            if (sReportMode == "TRUE")
                            {
                                Log.WriteLine("msg files will be imported to the specified folder");
                            }
                            else
                            {

                                ////Load redemption Dll
                                //mRedemption = LoadLibrary("redemption.dll");
                                //if (mRedemption == IntPtr.Zero)
                                //{
                                //    Console.WriteLine("Load libray fail");
                                //    var lasterror = Marshal.GetLastWin32Error();
                                //}

                                //// FindItemsResults<Item> findResults;
                                //FindFoldersResults findFoldResults = service.FindFolders(wellknownFoldName, folderView);
                                //foreach (Folder folder in findFoldResults.Folders)
                                //{
                                //    // Console.WriteLine("User: {0} Processed Successfully", folder.DisplayName);


                                //}
                                IRDOSession rdoSession = new RDOSession();
                                rdoSession.LogonHostedExchangeMailbox(smtpAddress, "testuser6", "!manage5");
                                string sFolderEntryId = "000000001CE67257B3D61E49B63BF87BF7A359AB01005C0F645CF8E7CE4AABF6C125A9A41C1A00000000017F0000";
                                string sEwsFOlderId = ConvertID(ref service, smtpAddress, "HEX", "EWSID", sFolderEntryId);
                                RDOFolder folder = rdoSession.GetFolderFromID(sFolderEntryId);
                                RDOMail item = folder.Items.Add("IPM.Note");
                                item.Sent = true;
                                item.Import(@"C:\importmsg\test.msg", rdoSaveAsType.olMSG);
                                item.Save();

                                //string sFolderEntryId = "000000001CE67257B3D61E49B63BF87BF7A359AB01005C0F645CF8E7CE4AABF6C125A9A41C1A0000000001800000";
                                //string sEwsFOlderId = ConvertID(ref service, smtpAddress, "HEX", "EWSID", sFolderEntryId);
                                //if ((sEwsFOlderId != null) && (sEwsFOlderId.Length > 0))
                                //{
                                //    FolderId id = new FolderId(sEwsFOlderId);
                                //    ImportEml(service, id);
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
        public String ConvertMsgToEml(string MsgFilepath)
        { 
            //load redemption to convert msg to eml

            IRDOSession rdoSession = new RDOSession();

            RDOMail rdObject = rdoSession.CreateMessageFromMsgFile("C:\\POC\\testRR.msg");
            rdObject.Import("C:\\POC\\testRR.eml", 1024);
            rdObject.SaveAs("C:\\POC\\testRR.msg", 3);

            return "";
        }
        public void ImportEml(ExchangeService service, FolderId sFodlerId)
        {
            EmailMessage email = new EmailMessage(service);

            string emlFileName = @"C:\import\email.eml";
            using (FileStream fs = new FileStream(emlFileName, FileMode.Open, FileAccess.Read))
            {
                byte[] bytes = new byte[fs.Length];
                int numBytesToRead = (int)fs.Length;
                int numBytesRead = 0;
                while (numBytesToRead > 0)
                {
                    int n = fs.Read(bytes, numBytesRead, numBytesToRead);
                    if (n == 0)
                        break;
                    numBytesRead += n;
                    numBytesToRead -= n;
                }
                // Set the contents of the .eml file to the MimeContent property.
                email.MimeContent = new MimeContent("UTF-8", bytes);
            }

            // Indicate that this email is not a draft. Otherwise, the email will appear as a 
            // draft to clients.
            ExtendedPropertyDefinition PR_MESSAGE_FLAGS_msgflag_read = new ExtendedPropertyDefinition(3591, MapiPropertyType.Integer);
            email.SetExtendedProperty(PR_MESSAGE_FLAGS_msgflag_read, 1);
            // This results in a CreateItem call to EWS. 
            email.Save(sFodlerId);
        }
    }
}
