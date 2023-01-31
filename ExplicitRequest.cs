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
    public class ExplicitRequest
    {
        private const string PATH_SEPARATOR = "/";

        public ExplicitRequest(IEMFilingRequest explicitRequest)
        {
            // Retrieve the failed filing request information pertinent to this application.
            sid = explicitRequest.SID;
            type = explicitRequest.RequestType;//.ConvertToExplicitRequestType();
            mailboxID = explicitRequest.Mailbox;
            userID = explicitRequest.UserID;
            projectFolderPath = ElaboratePath(explicitRequest.Folder.Path);
            submissionDate = explicitRequest.SubmissionDate;
            status = explicitRequest.StatusCode;
            statusDescription = explicitRequest.StatusMessage;
            databaseName = explicitRequest.Database.Name;
            retryCount = explicitRequest.RetryCount;
            exchFolderID = explicitRequest.EMFolder;

            if ((explicitRequest.RetryCount < 15) && (explicitRequest.StatusCode != EMRequestStatus.EMRequestFailure))
                isActive = true;
            else
                isActive = false;

            if (exchFolderID.Length > 0)
            {
                if (exchFolderID.Contains("EwsFolderId:"))
                    isInMappedFolder = true;
                else
                    isInMappedFolder = false;
            }
            else
                isInMappedFolder = false;
            //emailGUID = ExtractEmailGUID(explicitRequest);
        }

        public static string ElaboratePath(IManFolders pathToElaborate)
        {
            string workingPath = String.Empty;

            // The use of the last folder variable is intended to build the path of folders without the name (last element).
            foreach (IManFolder currentFolder in pathToElaborate)
            {
                workingPath += PATH_SEPARATOR + currentFolder.Name;
            }
            return workingPath;
        }

        /// <summary>
        /// ExtractEmailGUID returns a unique identifier for the email that this filing request represents.
        /// </summary>
        /// <param name="explicitRequest">the filing request from which the identifier should be retrieved</param>
        /// <returns>a unique identifier for the email associated with the filing request</returns>
        //private static string ExtractEmailGUID(IEMFilingRequest explicitRequest)
        //{
        //    string emailGUID = "";

        //    foreach (string otherPropertyString in explicitRequest.OtherProperties)
        //    {
        //        string[] otherPropertyPair = otherPropertyString.Split(DocumentManagementSystem.FILING_REQUEST_OTHER_PROPERTIES_SEPARATOR);
        //        if (otherPropertyPair.Length == 2)
        //        {
        //            string name = otherPropertyPair[0];
        //            string value = otherPropertyPair[1];
        //            if (name.Equals(DocumentManagementSystem.FILING_REQUEST_OTHER_PROPERTIES_EMAIL_GUID_NAME))
        //            {
        //                emailGUID = value;
        //            }
        //        }
        //    }

        //    return emailGUID;
        //}

        public int SID
        {
            get
            {
                return sid;
            }
        }

        public EMFilingRequestType Type
        {
            get
            {
                return type;
            }
        }

        public string MailboxID
        {
            get
            {
                return mailboxID;
            }
        }

        public string UserID
        {
            get
            {
                return userID;
            }
        }

        public string ProjectFolderPath
        {
            get
            {
                return projectFolderPath;
            }
        }

        public DateTime SubmissionDate
        {
            get
            {
                return submissionDate;
            }
        }

        public EMRequestStatus Status
        {
            get
            {
                return status;
            }
        }

        public string StatusDescription
        {
            get
            {
                return statusDescription;
            }
        }

        public string DatabaseName
        {
            get
            {
                return databaseName;
            }
        }

        public string EmailGUID
        {
            get
            {
                return emailGUID;
            }
        }

        public int RetryCount
        {
            get
            {
                return retryCount;
            }
        }

        public string ExchFolderID
        {
            get
            {
                return exchFolderID;
            }
        }
        public bool IsActive
        {
            get
            {
                return isActive;
            }
        }

        public bool IsInMappedFolder
        {
            get
            {
                return isInMappedFolder;
            }
        }
        private const int OTHER_PROPERTIES_ELEMENT_COUNT = 2;
        private const int OTHER_PROPERTIES_NAME_INDEX = 0;
        private const int OTHER_PROPERTIES_VALUE_INDEX = 1;

        private int sid;
        //private ExplicitRequestType type;
        private EMFilingRequestType type;
        private string mailboxID;
        private string userID;
        private string projectFolderPath;
        private DateTime submissionDate;
        private EMRequestStatus status;
        private string statusDescription;
        private string databaseName;
        private string emailGUID;
        private int retryCount;
        private string exchFolderID;
        private bool isActive;
        private bool isInMappedFolder;
    }
}
