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
    public class ExchangeQueuedEmails
    {
        private const string PATH_SEPARATOR = "/";

        public ExchangeQueuedEmails()
        {
            User = "";
            EmailId = "";
            subject = "";
            messageClass = "";
            messageId = "";
            entryId = "";
            ewsId = "";
            parentFolderEWSId = "";
            parentFolderEntryId = "";
            parentFolderName = "";
            sentDate = "";
            iExistInWorkServer = 2;
            lastModifiedTime = "";
            searchKey = "";
            PrjId = 0;
        }

        public string User;
        public string EmailId;
        public string subject;
        public string messageClass;
        public string messageId;
        public string entryId;
        public string ewsId;
        public string parentFolderEWSId;
        public string parentFolderEntryId;
        public string parentFolderName;
        public string sentDate;
        public int iExistInWorkServer;
        public string lastModifiedTime;
        public string searchKey;
        public int PrjId;

        public ExplicitRequest explicitReq;
        public FiledEmailDetails filedEmailDetails;
    }
}
