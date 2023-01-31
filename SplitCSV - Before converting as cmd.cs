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

        public bool SplitCSVForLinkedFolders(String CSVFilePath)
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
                String sFolderPath;
                String sUserId = "";
                String sUserSMTP;
                String sServer;
                String sEnabled;
                String sStatus;
                String sCurrentUser = "";
                String sData = "";
                String sFolderEntryId = "";

                while ((line = file.ReadLine()) != null)
                {
                    line.Trim();
                    if (String.IsNullOrEmpty(line))
                        continue;

                    String[] Tokens = line.Split(",".ToCharArray());

                    if (8 > Tokens.Length)
                    {
                        throw new Exception(String.Format("Invalid entry in {0} at line {1}", CSVFilePath, sLineNum));
                    }

                    sStatus = Tokens[4];
                    sEnabled = Tokens[3];
                    sFolderPath = Tokens[5]; // EM_REQUEST - FOLDER_PATH
                    sUserId = Tokens[0]; // USERID
                    sUserSMTP = Tokens[7].ToUpper(); // DOCUSER - EMAIL
                    sServer = Tokens[6].ToUpper(); // DOCUSER - EXCH_AUTO_DISC
                    sFolderEntryId = Tokens[2].ToUpper();
                    if (sStatus == "-6")
                        continue;

                    if (sEnabled != "Y")
                        continue;

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


                        fileWrite = new System.IO.StreamWriter("Output\\" + sUserId + ".csv", true);
                        fileBatchWrite = new System.IO.StreamWriter("Output\\bat\\" + sUserId + ".bat", false);
                        sServer = "WEBMAIL.FREDLAW.COM";
                        //sServer = "outlook_us.intfirm.com";

                        //sData = "call start /wait EWSTestApp.exe SCAN-FILED-QUEUED-EMAILS svc_efscloud K[kjgd036]K " + sUserSMTP + " " + sServer/*"\"\""*/ + " " + "c:\\CSV\\" + sUserId + ".csv" + " 2013-06-01 2015-09-11 True";
                        sData = "call start /wait EWSTestApp.exe SCAN-LINKED-FOLDERS svc_efscloud K[kjgd036]K " + sUserSMTP + " " + sServer/*"\"\""*/ + " " + "c:\\CSV\\" + sUserId + ".csv" + " 2013-06-01 2015-09-11 False";
                        //sData = "call start /wait EWSTestApp.exe SCAN-LINKED-FOLDERS ImpersonatorSMTP ImpersonatorPassword " + sUserSMTP + " " + sServer/*"\"\""*/ + " " + "c:\\CSV\\" + sUserId + ".csv" + " StartDate EndDate True";
                        fileBatchWrite.WriteLine(sData);
                        fileBatchWrite.WriteLine("call exit");

                        if (!m_oUserId.ContainsKey(sUserId))
                            m_oUserId.Add(sUserId, sUserId);
                    }

                    if (sFolderPath.Contains("EwsID:"))
                    {
                        String[] Toks = sFolderPath.Split(":".ToCharArray());
                        sFolderPath = Toks[1];
                    }
                    sData = sFolderEntryId + "," + sFolderPath + "," + sStatus + "," + sEnabled;
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
