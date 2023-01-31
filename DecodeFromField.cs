using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace EWSTestApp
{
    class DecodeFromField
    {
        string DecodeQuotedPrintables(ref string input, string charSet)
        {
            if (string.IsNullOrEmpty(charSet))
            {
                var charSetOccurences = new Regex(@"=\?.*\?Q\?", RegexOptions.IgnoreCase);
                var charSetMatches = charSetOccurences.Matches(input);
                foreach (Match match in charSetMatches)
                {
                    charSet = match.Groups[0].Value.Replace("=?", "").Replace("?Q?", "");
                    input = input.Replace(match.Groups[0].Value, "").Replace("?=", "");
                }
            }

            Encoding enc = new UTF8Encoding();
            // Encoding enc = new ASCIIEncoding();
            if (!string.IsNullOrEmpty(charSet))
            {
                try
                {
                    enc = Encoding.GetEncoding(charSet);
                }
                catch
                {
                    enc = new ASCIIEncoding();
                }
            }

            //decode iso-8859-[0-9]
            var occurences = new Regex(@"=[0-9A-Z]{2}", RegexOptions.Multiline);
            var matches = occurences.Matches(input);
            foreach (Match match in matches)
            {
                try
                {
                    byte[] b = new byte[] { byte.Parse(match.Groups[0].Value.Substring(1), System.Globalization.NumberStyles.AllowHexSpecifier) };
                    char[] hexChar = enc.GetChars(b);
                    input = input.Replace(match.Groups[0].Value, hexChar[0].ToString());
                    //input = input.Replace("=?iso-8859-1?Q?"," ");
                    //input = input.Replace("_", " ").Replace("?=", " ");
                }
                catch
                { ;}
            }

            //decode base64String (utf-8?B?)
            occurences = new Regex(@"\?utf-8\?B\?.*\?", RegexOptions.IgnoreCase);
            //if ((input.Length % 4) != 0)
            //{
            //    input = input.Remove(input.Length - 1);
            //}
           // input = input.Replace("=\r\n", "");
            matches = occurences.Matches(input);
            foreach (Match match in matches)
            {
                if (((match.Groups[0].Value.Replace("?utf-8?B?", "").Replace("?UTF-8?B?", "").Replace("?","")).Length) % 4 != 0)
                {
                    byte[] b = Convert.FromBase64String(match.Groups[0].Value.Replace("?utf-8?B?", "").Replace("?UTF-8?B?", "").Replace("?", "="));
                    string temp = Encoding.UTF8.GetString(b);
                    input = input.Replace(match.Groups[0].Value, temp);
                }
                else {
                    byte[] b = Convert.FromBase64String(match.Groups[0].Value.Replace("?utf-8?B?", "").Replace("?UTF-8?B?", "").Replace("?", ""));
                    string temp = Encoding.UTF8.GetString(b);
                    input = input.Replace(match.Groups[0].Value, temp);
                }
            }

            input = input.Replace("=\r\n", "").Replace("_", " ").Replace("=","");

            return input;
        }
        public void ReadCsvFileToDecode(string inputfile)
        {
            Int32 iDocId = 0;
            String sFrom;
            StreamWriter OuputCSV = null;

            System.IO.StreamReader file = new System.IO.StreamReader(inputfile);
            if (File.Exists("DecodedFromField.csv"))
            {
                File.Delete("Decodedfromfield.csv");
                OuputCSV = new StreamWriter("DecodedFromField.csv", true);
                OuputCSV.AutoFlush = true;
            }
            else
            {
                OuputCSV = new StreamWriter("DecodedFromField.csv", true);
                OuputCSV.AutoFlush = true;
            }

            string line;
            //while ((line = file.ReadLine()) != null)
            while(!file.EndOfStream)
            {
                line = file.ReadLine();
                line.Trim();
                if(line == ""){
                continue;
                }
                String[] Tokens = line.Split(",".ToCharArray());
                if (Tokens[0].ToUpper() == "DOCNUM" && Tokens[1].ToUpper() == "C13ALIAS")
                {
                    continue;
                }
                iDocId = Convert.ToInt32((Tokens[0].ToUpper()));
                sFrom = Tokens[1];
                sFrom = sFrom.Replace("\"", "");
                DecodeFromField obj = new DecodeFromField();
                obj.DecodeQuotedPrintables(ref sFrom, null);
                OuputCSV.WriteLine("{0},{1}", iDocId, sFrom);
            }
            file.Close();
            OuputCSV.Close();

        }
    }


}
