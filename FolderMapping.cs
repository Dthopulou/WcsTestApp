using Com.Interwoven.WorkSite.iManage;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWSTestApp
{
   
    public class FolderMapping
    {
        private const string PATH_SEPARATOR = "/";

        public FolderMapping(IEMFolderMapping folderMapping)
        {
            // Retrieve the failed filing request information pertinent to this application.
            sid = folderMapping.SID;
            type = folderMapping.RequestType;//.ConvertToExplicitRequestType();
            mailboxID = folderMapping.Mailbox;
            userID = folderMapping.UserID;
            projectFolderPath = ElaboratePath(folderMapping.Folder.Path);            
            status = folderMapping.StatusCode;            
            statusDescription = folderMapping.StatusMessage;
            databaseName = folderMapping.Database.Name;           
            exchFolderID = folderMapping.EMFolder;
            prjID = folderMapping.FolderID;
            lastSyncTime = folderMapping.LastSync.ToString();
            isActive = folderMapping.Enabled;
            foldEntryId = folderMapping.EMFolder;
            OtherProperties = folderMapping.OtherProperties;

            IEMFolderMapping2 fold2 = (IEMFolderMapping2) folderMapping;
            if (fold2 != null)
            {
                sOperator = fold2.Operator;
            }
            
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

     
        public int sid;
        //private ExplicitRequestType type;
        public EMFilingRequestType type;
        public string mailboxID;
        public string userID;
        public string projectFolderPath;
        public DateTime submissionDate;
        public EMFolderMappingStatus status;
        public string statusDescription;
        public string databaseName;
        public string emailGUID;
        public int retryCount;
        public string exchFolderID;
        public bool isActive;
        public bool isInMappedFolder;
        public int prjID;

        public string sOperator;
        public string lastSyncTime;
        public string foldEntryId;
        public ManStrings OtherProperties;
        
    }
}
