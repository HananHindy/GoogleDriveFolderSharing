using Google.Apis.Drive.v2.Data;
using System;
using System.Collections.Generic;

namespace GoogleDriveManagement
{
    class Program
    {
        struct Group
        {
            public String Email;
            public String TeamName;
            public String Department;
        }

        private static List<Group> groups = new List<Group>();

        static void Main(string[] args)
        {
            String sheetName = Configurations.FileName;
            ReadExcelSheet(sheetName);

            GoogleDriveHelper google = new GoogleDriveHelper();
            String parentfolderName = Configurations.RootFolderName;

            foreach (Group group in groups)
            {
                string folderName = parentfolderName;
                if (!string.IsNullOrEmpty(group.Department))
                {
                    folderName += '/' + group.Department.ToUpper();
                }

                File file = google.CreateDirectory(folderName, group.TeamName);
                String folderId = file.Id;
                google.Share(folderId, group.Email, "user", "reader");
                google.Share(folderId, group.Email, "user", "writer");
            }
        }

        static void ReadExcelSheet(String sheetPath)
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range range;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(sheetPath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            int rw = range.Rows.Count;
            int cl = range.Columns.Count;

            for (int rCnt = 2; rCnt <= rw; rCnt++)
            {
                Group student;
                double value = ((range.Cells[rCnt, 1] as Microsoft.Office.Interop.Excel.Range).Value2);
                student.TeamName = value.ToString();
                string email = (string)(range.Cells[rCnt, 2] as Microsoft.Office.Interop.Excel.Range).Value2;
                student.Email = email;
                student.Department = ((string)(range.Cells[rCnt, 3] as Microsoft.Office.Interop.Excel.Range).Value2);
                groups.Add(student);
            }


            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
        }
    }
}
