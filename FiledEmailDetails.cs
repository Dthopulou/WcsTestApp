using Com.Interwoven.WorkSite.iManage;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWSTestApp
{
    /// <summary>
    /// ExplicitRequest is this application's view of a filing request, based largely on the WorkSite object model's view.
    /// </summary>
    public class FiledEmailDetails
    {
        private const string PATH_SEPARATOR = "/";

        public FiledEmailDetails()
        {
            messageId = "";
            DocNum = 0;
            Version = 0;
        }

        public string messageId;
        public double DocNum;
        public int Version;
    }
}
