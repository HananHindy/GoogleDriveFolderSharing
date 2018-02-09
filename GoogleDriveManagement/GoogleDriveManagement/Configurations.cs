using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace GoogleDriveManagement
{
    public class Configurations
    {
        public static string AuthFilePath
        {
            get
            {
                return ConfigurationManager.AppSettings["AuthFilePath"];
            }
        }

        public static string FileName
        {
            get
            {
                return ConfigurationManager.AppSettings["FileName"];
            }
        }

        public static string RootFolderName
        {
            get
            {
                return ConfigurationManager.AppSettings["RootFolderName"];
            }
        }
    }
}
