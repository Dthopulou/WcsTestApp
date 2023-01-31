using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace EWSTestApp
{
    class SplitCSV
    {
        
        private Dictionary<String, String> m_oUserId = null;

        public bool GenerateBatchFileToResetEmailsOnEntireMailbox(string[] args)
        {
            bool bRet = false;

            String InputUsersFile = args[1];

            do
            {
                if (args.Length < 9)
                {
                    Console.WriteLine("Syntax: <Command> <InputUsersFile> <ImpersonatorSMTP> <password> <exchange server name> <RootFolder> <QueuedOnly> <Start Date> <End Date> <Report Mode>");
                    Console.WriteLine("Example: GEN_BAT_ENTIRE_MB UsersFile ImpersonatorSMTPAddress@dev.local password xchange.dev.local 1 True 2015-06-19 2015-06-20 True");
                    return bRet;
                }

                string line;
                String sLocOfOutputCSV = args[1];
                System.IO.StreamReader file = new System.IO.StreamReader(InputUsersFile);
                System.IO.StreamWriter fileGlobalBatchWrite = null;
                Int32 sLineNum = 0;
                String sUserName;
                String sUserEmailAddr = "";

                if (!File.Exists(InputUsersFile))
                {
                    break;
                }

                bool bEchoOff = false;
                String sData = "";
                while ((line = file.ReadLine()) != null)
                {
                    line.Trim();
                    if (String.IsNullOrEmpty(line))
                        continue;

                    String[] Tokens = line.Split(",".ToCharArray());

                    if (2 > Tokens.Length)
                    {
                        Console.WriteLine(String.Format("Invalid entry in {0} at line {1}", InputUsersFile, sLineNum));
                        continue;
                    }
                    sUserName = Tokens[0];
                    sUserEmailAddr = Tokens[1];

                    if ((sUserName == "") || (sUserEmailAddr == ""))
                    {
                        Console.WriteLine("Username or User email address is null");
                        continue;
                    }

                    if (fileGlobalBatchWrite == null)
                        fileGlobalBatchWrite = new System.IO.StreamWriter("ResetEmails.bat", true);

                    if (!bEchoOff)
                    {
                        fileGlobalBatchWrite.WriteLine("@echo off");
                        bEchoOff = true;
                    }

                    sData = "call start /wait EWSTestApp.exe SCAN-MAILBOX-RESET-QUEUED-EMAILS " + args[2] + " " + args[3] + " " + sUserEmailAddr + " " + "\"" + args[4]/*"\"\""*/ + "\"" + " " + args[5] + " " + args[6] + " " + args[7] + " " + args[8] + " " + args[9];

                    fileGlobalBatchWrite.WriteLine("echo " + sUserName);
                    fileGlobalBatchWrite.WriteLine(sData);
                    fileGlobalBatchWrite.WriteLine("call start exit");

                }
                if (fileGlobalBatchWrite != null)
                {
                    fileGlobalBatchWrite.Close();
                    fileGlobalBatchWrite = null;
                }
            } while (false);

            return true;
        }
        public bool SplitCSVForLinkedFolders(string[] args)//String CSVInputFilePath)
        {
            bool bRet = false;
            String sEnabled;
            if (args.Length < 10)
            {
                Console.WriteLine("Syntax: <Command> <InputCSVFile> <ImpersonatorSMTP> <password> <exchange server name> <OutputCSVFilePath> <Start Date> <End Date> <CaptureDisabledMapping> <Report Mode>");
                Console.WriteLine("Example: SPLITXML InputEM_REQ_CSV_File ImpersonatorSMTPAddress@dev.local password xchange.dev.local Output_CSV_Directory_Path 2015-06-19 2015-06-20 True False");
                return bRet;
            }

            String CSVInputFilePath = args[1];
            String sImpersonatorSTMPAdd = args[2];
            String sPassword = args[3];
            String sServer = args[4];
            String sLocOfOutputCSV = args[5];
            String sStartDt = args[6];
            String sEndDt = args[7];
            String sCaptureDisabledMapping = args[8];
            String sReportMode = args[9];

            //SCAN-LINKED-FOLDERS admin@imanage.microsoftonline.com !Manage.2015 jsmith@imanage.microsoftonline.com ch1prd0410.outlook.com JSMITH.csv 2013-08-01 2015-09-15 False
            //SCAN-LINKED-FOLDERS svc_efscloud K[kjgd036]K " + sUserSMTP + " " + sServer/*"\"\""*/ + " " + "c:\\CSV\\" + sUserId + ".csv" + " 2013-06-01 2015-09-11 False";

            //"call start /wait EWSTestApp.exe SCAN-LINKED-FOLDERS " + sImpersonatorSTMPAdd + " " + sPassword + " " + sUserSMTP + " " + sServer/*"\"\""*/ + " " + sLocOfCSV + "\\" + sUserId + ".csv " + sStartDt + " " + sEndDt + " " + sReportMode;
            do
            {
                String bEnabled;
                

                if (!File.Exists(CSVInputFilePath))
                {
                    break;
                }
                m_oUserId = new Dictionary<String, String>();
                System.IO.StreamReader file = new System.IO.StreamReader(CSVInputFilePath);
                System.IO.StreamWriter fileWrite = null;
                System.IO.StreamWriter fileBatchWrite = null;
                System.IO.StreamWriter fileGlobalBatchWrite = null;

                string line;

                Int32 sLineNum = 0;
                String sFolderPath;
                String sUserId = "";
                String sUserSMTP;
                
                
                String sStatus;
                String sCurrentUser = "";
                String sData = "";
                String sFolderEntryId = "";
                String sBatchFileLoc = "";

                while ((line = file.ReadLine()) != null)
                {
                    line.Trim();
                    if (String.IsNullOrEmpty(line))
                        continue;

                    String[] Tokens = line.Split(",".ToCharArray());

                    if (8 > Tokens.Length)
                    {
                        throw new Exception(String.Format("Invalid entry in {0} at line {1}", CSVInputFilePath, sLineNum));
                    }

                    sStatus = Tokens[4];
                    sEnabled = Tokens[3];
                    sFolderPath = Tokens[5]; // EM_REQUEST - FOLDER_PATH
                    sUserId = Tokens[0]; // USERID
                    sUserSMTP = Tokens[7].ToUpper(); // DOCUSER - EMAIL
                    if (sServer.Length == 0)
                        sServer = Tokens[6].ToUpper(); // DOCUSER - EXCH_AUTO_DISC
                    sFolderEntryId = Tokens[2].ToUpper();

                    int iPos = sServer.IndexOf('>');

                    if (iPos > 0)
                        sServer = sServer.Substring(iPos+1, sServer.Length - iPos - 1);

                    if (sStatus == "-6")
                        continue;

                    

                    int iMatch = String.Compare(sCaptureDisabledMapping, "True", false);
                    if ((iMatch == 0) || (iMatch == -1))
                    {
                        //continue;
                    }
                    else
                    {
                        if (sEnabled != "Y")
                            continue;
                    
                    }

                    //if (sEnabled != "Y")
                      //  continue;


                    if (sCurrentUser != sUserId)
                    {
                        sCurrentUser = sUserId;
                        if (fileWrite != null)
                        {
                            fileWrite.Close();
                            fileWrite = null;
                        }

                        if (fileBatchWrite != null)
                        {
                            fileBatchWrite.Close();
                            fileBatchWrite = null;
                        }

                        if (!Directory.Exists(sLocOfOutputCSV))//"Output"))
                            Directory.CreateDirectory(sLocOfOutputCSV);//"Output");

                        sBatchFileLoc = sLocOfOutputCSV;
                        //sBatchFileLoc += "\\bat";

                       // if (!Directory.Exists(sBatchFileLoc))//("Output\\bat"))
                        //    Directory.CreateDirectory(sBatchFileLoc);// ("Output\\bat");

                        fileWrite = new System.IO.StreamWriter(sLocOfOutputCSV + "\\" + sUserId + ".csv", true);
                        fileBatchWrite = new System.IO.StreamWriter(sBatchFileLoc + "\\" + sUserId + ".bat", false);
                        //sServer = "WEBMAIL.FREDLAW.COM";
                        //sServer = "outlook_us.intfirm.com";

                        sData = "call start /wait EWSTestApp.exe SCAN-LINKED-FOLDERS " + sImpersonatorSTMPAdd + " " + sPassword + " " + sUserSMTP + " " + sServer/*"\"\""*/ + " " + sLocOfOutputCSV + "\\" + sUserId + ".csv " + sStartDt + " " + sEndDt + " " + sReportMode;

                        //sData = "call start /wait EWSTestApp.exe SCAN-FILED-QUEUED-EMAILS svc_efscloud K[kjgd036]K " + sUserSMTP + " " + sServer/*"\"\""*/ + " " + "c:\\CSV\\" + sUserId + ".csv" + " 2013-06-01 2015-09-11 True";
                        //sData = "call start /wait EWSTestApp.exe SCAN-LINKED-FOLDERS svc_efscloud K[kjgd036]K " + sUserSMTP + " " + sServer/*"\"\""*/ + " " + "c:\\CSV\\" + sUserId + ".csv" + " 2013-06-01 2015-09-11 False";
                        //sData = "call start /wait EWSTestApp.exe SCAN-LINKED-FOLDERS ImpersonatorSMTP ImpersonatorPassword " + sUserSMTP + " " + sServer/*"\"\""*/ + " " + "c:\\CSV\\" + sUserId + ".csv" + " StartDate EndDate True";
                        fileBatchWrite.WriteLine(sData);
                        fileBatchWrite.WriteLine("call exit");

                        if (!m_oUserId.ContainsKey(sUserId))
                            m_oUserId.Add(sUserId, sUserId);
                    }

                    //if (sFolderPath.Contains("EwsID:"))
                    //{
                    //    String[] Toks = sFolderPath.Split(":".ToCharArray());
                    //    sFolderPath = Toks[1];
                    //}
                    //sData = sFolderEntryId + "," + sFolderPath + "," + sStatus + "," + sEnabled;
                    sData = sFolderEntryId + "," + sFolderEntryId + "," + sStatus + "," + sEnabled;
                    fileWrite.WriteLine(sData);
                }

                sCurrentUser = sUserId;
                if (fileWrite != null)
                {
                    fileWrite.Close();
                    fileWrite = null;
                }

                if (fileBatchWrite != null)
                {
                    fileBatchWrite.Close();
                    fileBatchWrite = null;
                }

                bool bEchoOff = false;
                fileGlobalBatchWrite = new System.IO.StreamWriter(sBatchFileLoc + "\\" + "ResetEmails.bat", true);
                foreach (KeyValuePair<String, String> Entry in m_oUserId)
                {
                    if (!bEchoOff)
                    {
                        fileGlobalBatchWrite.WriteLine("@echo off");
                        bEchoOff = true;
                    }

                    fileGlobalBatchWrite.WriteLine("echo " + Entry.Key);
                    fileGlobalBatchWrite.WriteLine("call start /wait " + Entry.Key + ".bat");
                    fileGlobalBatchWrite.WriteLine("call start exit");

                }

                if (fileGlobalBatchWrite != null)
                {
                    fileGlobalBatchWrite.Close();
                    fileGlobalBatchWrite = null;
                }
                
            } while (false);
            return true;
        }

        public bool SplitCSVFileWithValidEntryIdOrGuid(String CSVFilePath)
        {
            bool bRet = false;
            do
            {
                if (!File.Exists(CSVFilePath))
                {
                    break;
                }
                //m_oUserId = new Dictionary<String, String>();

                System.IO.StreamReader file = new System.IO.StreamReader(CSVFilePath);
                System.IO.StreamWriter fileWrite = null;
                System.IO.StreamWriter fileBatchWrite = null;
                //System.IO.StreamWriter fileGlobalBatchWrite = null;

                string line;

                Int32 sLineNum = 0;
                String sEntryId;
                String sFolderPath;
                String sUserSMTP;
                String sServer;
                String sEmailGuid;
                String sStatus;
                String sUserId = "";
                String sCurrentUser = "";
                String sData = "";

                while ((line = file.ReadLine()) != null)
                {
                    line.Trim();
                    if (String.IsNullOrEmpty(line))
                        continue;

                    String[] Tokens = line.Split(",".ToCharArray());
                    if (5 > Tokens.Length)
                    {
                        throw new Exception(String.Format("Invalid entry in {0} at line {1}", CSVFilePath, sLineNum));
                    }

                    sEntryId = Tokens[0].ToUpper(); // EM_REQUEST - MSG_ID
                    sEmailGuid = Tokens[1].ToUpper(); // EM_REQUEST - EMAIL_GUID
                    sStatus = Tokens[2]; // Status
                    sFolderPath = Tokens[4]; // EM_REQUEST - FOLDER_PATH
                    sUserId = Tokens[5]; // USERID
                    sUserSMTP = Tokens[6].ToUpper(); // DOCUSER - EMAIL
                    sServer = Tokens[7].ToUpper(); // DOCUSER - EXCH_AUTO_DISC

                    if ((sUserId == "") || 
                        (sUserId == "NULL") || 
                        (sUserId == "null"))
                        continue;

                    if ((sEntryId.Length <= 4) && (sEmailGuid.Length <= 4))
                        continue;

                    if (sStatus != "-6")
                        continue;

                    if (sUserId.Contains(":"))
                        continue;

                    if (sCurrentUser != sUserId)
                    {
                        //if (m_oUserId.ContainsKey(sUserId))
                        //    continue;
                        //else
                        //{
                        //m_oUserId.Add(sUserId, sUserId);
                        sCurrentUser = sUserId;
                        if (fileWrite != null)
                        {
                            fileWrite.Close();
                            fileWrite = null;
                        }

                        if (fileBatchWrite != null)
                        {
                            fileBatchWrite.Close();
                            fileBatchWrite = null;
                        }


                        fileWrite = new System.IO.StreamWriter("Output\\" + sUserId + ".csv", true);
                        fileBatchWrite = new System.IO.StreamWriter("Output\\bat\\" + sUserId + ".bat", false);

                        sData = "call start /wait EWSTestApp.exe SCAN-FOLDERS svc_efscloud K[kjgd036]K " + sUserSMTP + " " + "\"\"" + " " + "c:\\CSV\\" + sUserId + ".csv" + " True";
                        fileBatchWrite.WriteLine(sData);
                        fileBatchWrite.WriteLine("call exit");

                        //if (!m_oUserId.ContainsKey(sUserId))
                        //    m_oUserId.Add(sUserId, sUserId);

                        //}
                    }

                    sData = sEntryId + "," + sEmailGuid + "," + sFolderPath + "," + sUserSMTP + "," + sServer;
                    fileWrite.WriteLine(sData);





                }

                sCurrentUser = sUserId;
                if (fileWrite != null)
                {
                    fileWrite.Close();
                    fileWrite = null;
                }

                if (fileBatchWrite != null)
                {
                    fileBatchWrite.Close();
                    fileBatchWrite = null;
                }

                //bool bEchoOff = false;
                //fileGlobalBatchWrite = new System.IO.StreamWriter("Output\\bat\\ResetEmails.bat", true);
                //foreach (KeyValuePair<String, String> Entry in m_oUserId)
                //{
                //    if (!bEchoOff)
                //    {
                //        fileGlobalBatchWrite.WriteLine("@echo off");
                //        bEchoOff = true;
                //    }

                //    fileGlobalBatchWrite.WriteLine("echo " + Entry.Key);
                //    fileGlobalBatchWrite.WriteLine("call start /wait " + Entry.Key + ".bat");
                //    fileGlobalBatchWrite.WriteLine("call start exit");

                //}

                //if (fileGlobalBatchWrite != null)
                //{
                //    fileGlobalBatchWrite.Close();
                //    fileGlobalBatchWrite = null;
                //}
                bRet = false;
            } while (false);
            return bRet;
        }

        public bool SplitCSVFile(String CSVFilePath)
        {
            bool bRet = false;
            do
            {
                if (!File.Exists(CSVFilePath))
                {
                    break;
                }
                m_oUserId = new Dictionary<String, String>();

                System.IO.StreamReader file = new System.IO.StreamReader(CSVFilePath);
                System.IO.StreamWriter fileWrite = null;
                System.IO.StreamWriter fileBatchWrite = null;
                System.IO.StreamWriter fileGlobalBatchWrite = null;

                string line;

                Int32 sLineNum = 0;
                String sEntryId;
                String sFolderPath;
                String sUserSMTP;
                String sServer;
                String sEmailGuid;
                String sUserId = "";
                String sCurrentUser = "";
                String sData = "";

                while ((line = file.ReadLine()) != null)
                {
                    line.Trim();
                    if (String.IsNullOrEmpty(line))
                        continue;

                    String[] Tokens = line.Split(",".ToCharArray());
                    if (5 > Tokens.Length)
                    {
                        throw new Exception(String.Format("Invalid entry in {0} at line {1}", CSVFilePath, sLineNum));
                    }

                    sEntryId = Tokens[0].ToUpper(); // EM_REQUEST - MSG_ID
                    sEmailGuid = Tokens[1].ToUpper(); // EM_REQUEST - EMAIL_GUID
                    sFolderPath = Tokens[2]; // EM_REQUEST - FOLDER_PATH
                    sUserId = Tokens[3];
                    sUserSMTP = Tokens[4].ToUpper(); // DOCUSER - EMAIL
                    sServer = Tokens[5].ToUpper(); // DOCUSER - EXCH_AUTO_DISC

                    if ((sUserId == "") || (sUserId == "NULL") || (sUserId == "null"))
                        continue;

                    if (sCurrentUser != sUserId)
                    {
                        //if (m_oUserId.ContainsKey(sUserId))
                        //    continue;
                        //else
                        //{
                            //m_oUserId.Add(sUserId, sUserId);
                            sCurrentUser = sUserId;
                            if (fileWrite  != null)
                            {
                                fileWrite.Close();
                                fileWrite = null;                                
                            }

                            if (fileBatchWrite != null)
                            {
                                fileBatchWrite.Close();
                                fileBatchWrite = null;
                            }


                            fileWrite = new System.IO.StreamWriter("Output\\"+sUserId+".csv",true);
                            fileBatchWrite = new System.IO.StreamWriter("Output\\bat\\" + sUserId + ".bat", false);

                            sData = "call start /wait EWSTestApp.exe SCAN-FOLDERS svc_efscloud K[kjgd036]K " + sUserSMTP + " " + "\"\"" + " " + "c:\\CSV\\" + sUserId + ".csv" + " 2015-03-05 True";
                            fileBatchWrite.WriteLine(sData);
                            fileBatchWrite.WriteLine("call exit");

                            if (!m_oUserId.ContainsKey(sUserId))
                                m_oUserId.Add(sUserId, sUserId);

                        //}
                    }

                    sData = sEntryId + "," + sEmailGuid + "," + sFolderPath + "," + sUserSMTP + "," + sServer;
                    fileWrite.WriteLine(sData); 

                    

                   
                    
                }

                sCurrentUser = sUserId;
                if (fileWrite != null)
                {
                    fileWrite.Close();
                    fileWrite = null;
                }

                if (fileBatchWrite != null)
                {
                    fileBatchWrite.Close();
                    fileBatchWrite = null;
                }

                bool bEchoOff = false;
                fileGlobalBatchWrite = new System.IO.StreamWriter("Output\\bat\\ResetEmails.bat", true);
                foreach (KeyValuePair<String, String> Entry in m_oUserId)
                {
                    if (!bEchoOff)
                    {
                        fileGlobalBatchWrite.WriteLine("@echo off");
                        bEchoOff = true;
                    }

                    fileGlobalBatchWrite.WriteLine("echo " + Entry.Key);
                    fileGlobalBatchWrite.WriteLine("call start /wait " + Entry.Key +".bat");
                    fileGlobalBatchWrite.WriteLine("call start exit");

                }
                
                if (fileGlobalBatchWrite != null)
                {
                    fileGlobalBatchWrite.Close();
                    fileGlobalBatchWrite = null;
                }
                bRet = false;
            } while (false);
            return bRet;
         }

    }
}
