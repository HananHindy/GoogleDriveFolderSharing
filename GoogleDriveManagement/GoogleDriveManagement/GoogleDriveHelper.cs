using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v2;
using Google.Apis.Drive.v2.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Threading;
using Google.Apis.Requests;

namespace GoogleDriveManagement
{
    public class GoogleDriveHelper
    {
        private string[] Scopes = {
                                      DriveService.Scope.Drive,
                                      DriveService.Scope.DriveFile,
                                      DriveService.Scope.DriveMetadata
                                  };

        private const string ApplicationName = "google-drive-management";

        private UserCredential _userCredential;
        public DriveService _driveService;

        public GoogleDriveHelper()
        {
            GetAuth();
        }

        public Google.Apis.Drive.v2.Data.File CreateDirectory(string parent_path, string file_name)
        {
            Google.Apis.Drive.v2.Data.File body = new Google.Apis.Drive.v2.Data.File();
            body.Title = file_name;
            body.MimeType = "application/vnd.google-apps.folder";
            body.Parents = new List<ParentReference>()
            {
                new ParentReference() { Id = GetIdByPath(parent_path) }
            };

            return _driveService.Files.Insert(body).Execute();
        }

        public void GetAuth()
        {
            string authFilePath = Configurations.AuthFilePath;

            using (var stream = new FileStream(authFilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
            {
                var credentials = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None);

                _userCredential = credentials.Result;
                if (credentials.IsCanceled || credentials.IsFaulted)
                    throw new Exception("cannot connect");

                _driveService = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = _userCredential,
                    ApplicationName = ApplicationName,
                });
            }
        }

        public void Share(string fileId, string value, string type, string role)
        {
            var permission = new Permission
            {
                Value = value,
                Type = type,
                Role = role,

            };

            var request = _driveService.Permissions.Insert(permission, fileId);
            request.EmailMessage = Configurations.SharingMessage;
            request.Execute();
        }

        public string GetIdByPath(string path)
        {
            if (path == "/" || path == "")
                return "root";

            string[] path_list = path.Split('/');
            List<Google.Apis.Drive.v2.Data.File> file_search_list = new List<Google.Apis.Drive.v2.Data.File>();

            FilesResource.ListRequest req = _driveService.Files.List();
            do
            {
                req.Q = "title='" + path_list.Last<string>() + "'";
                FileList file_search = req.Execute();
                file_search_list.AddRange(file_search.Items);
            } while (!String.IsNullOrEmpty(req.PageToken));

            if (file_search_list.Count == 1)
            {
                return file_search_list.First<Google.Apis.Drive.v2.Data.File>().Id;
            }
            else
            {
                int last_idx = path_list.Length - 2;
                Google.Apis.Drive.v2.Data.File ret = new Google.Apis.Drive.v2.Data.File();
                foreach (Google.Apis.Drive.v2.Data.File f in file_search_list)
                {
                    if (GetRecursiveParent(path, f.Parents, last_idx) != null)
                    {
                        ret = f;
                        break;
                    }
                }

                return ret.Id;
            }
        }

        public string GetRecursiveParent(string path, IList<ParentReference> parent, int idx)
        {
            string[] path_list = path.Split('/');
            FilesResource.GetRequest parent_req = _driveService.Files.Get(parent[0].Id);
            Google.Apis.Drive.v2.Data.File parent_file = parent_req.Execute();

            if (parent[0].IsRoot.Value)
                return parent[0].Id;
            if (parent_file.Title == path_list[idx] && !parent[0].IsRoot.Value)
                return GetRecursiveParent(path, parent_file.Parents, idx - 1);
            else
                return null;
        }
    }
}
