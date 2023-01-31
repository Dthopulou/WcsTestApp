using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace EWSTestApp
{
    class EmailStatus
    {
        StreamWriter Log = new StreamWriter("EWSEmailScanLog.txt", true);

        public void Execute(string[] args)
        {
            try
            {
                do
                {
                    Log.AutoFlush = true;

                    if (args.Length < 7)
                    {
                        Console.WriteLine("");
                        Console.WriteLine("FAILED!!! - Invalid parameters");
                        Console.WriteLine("");
                        Console.WriteLine("Syntax: <Command> <ImpersonatorSMTP> <password> <endUserSMTP> <exchange server name> <User EM_REQUEST CSV filePath> <Start date> <RunReportMode>");
                        //SCAN-FOLDERS admin2@imanage.microsoftonline.com !wov2014 jsmith@imanage.microsoftonline.com ch1prd0410.outlook.com d:\Resubmit1.csv 2015-03-16
                        Console.WriteLine("Example: SCAN-EMAILS ImpersonatorSMTPAddress@dev.local password endUserSMTPAddress xchange.dev.local c:\\User.csv True");

                        break;
                    }
                } while (false);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
            }
            finally
            {

            }
        }
    }
}
