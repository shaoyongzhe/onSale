using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
namespace ExcelWorkbook4
{
    class pub
    {
        public int startindex = 10;
        public static void InsertLinkValue(WorksheetBase sheet1, 
                                            string cell, 
                                            string hyperlinkaddress)
        {
            Microsoft.Office.Interop.Excel.Range rangeToHoldHyperlink = sheet1.get_Range(cell);
            string hyperlinkTargetAddress = hyperlinkaddress;

            sheet1.Hyperlinks.Add(
                rangeToHoldHyperlink,
                hyperlinkaddress,
                string.Empty,
                "戳我，打开文件夹",
                hyperlinkaddress);
        }

        public static void ClearContent(WorksheetBase sheet1, string start_cell, string end_cell)
        {
            Excel.Range rngClear = sheet1.get_Range(string.Format("{0}:{1}", start_cell, end_cell));
            rngClear.Clear();

        }
        public static bool reloadDirectory(WorksheetBase sheet1)
        {
            DialogResult result1 = MessageBox.Show("是否确定要重新更新数据?（更新之后，原来的数据将全部清除）",
                                                 "警告！！！",
                                                  MessageBoxButtons.YesNo);
            if (result1 == DialogResult.No)
            {
                return true;
            }

            string timeformat = "yyyy-MM-dd HH-mm-ss,fff";
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            folderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyComputer;
            DialogResult result = folderBrowserDialog1.ShowDialog();

            if (result == DialogResult.OK)
            {
                InsertLinkValue(sheet1, "A3", folderBrowserDialog1.SelectedPath);
                ClearContent(sheet1, "A10", "Z100");
                
                string[] files = Directory.GetDirectories(folderBrowserDialog1.SelectedPath);

                int index = 8;
                Excel.Range rng = sheet1.get_Range(string.Format("A{0}:C{0}", index));
                foreach (var fileitem in files.OrderBy(x => x))
                {
                    string filename = fileitem.Split('\\').Last();
                    DateTime creationTime = File.GetCreationTime(fileitem);
                    DateTime lastWriteTime = File.GetLastWriteTime(fileitem);
                    string[] fileattributes = {filename, creationTime.ToString(timeformat),
                                                lastWriteTime.ToString(timeformat)};

                    rng.set_Value(Missing.Value, fileattributes);

                    InsertLinkValue(sheet1, string.Format("D{0}", index), fileitem);
                    index += 1;
                    rng = sheet1.get_Range(string.Format("A{0}:C{0}", index));
                }

                MessageBox.Show("共有文件夹总数: " + files.Length.ToString(), "Message");
            }
            return true;
        }

        public static Dictionary<string,List<string>> GetFiles(string rootDir)
        {
            string timeformat = "yyyy-MM-dd HH-mm-ss,fff";
            Dictionary<string, List<string>> fileDict = new Dictionary<string, List<string>>();
            string[] files = Directory.GetDirectories(rootDir);
            foreach(var fileitem in files.OrderBy(x => x))
            {
                string filename = fileitem.Split('\\').Last();
                DateTime creationTime = File.GetCreationTime(fileitem);
                DateTime lastWriteTime = File.GetLastWriteTime(fileitem);
                fileDict[filename] = new List<string> {filename, creationTime.ToString(timeformat),
                                                lastWriteTime.ToString(timeformat), fileitem};
            }

            return fileDict;
        }

        public static Dictionary<string, int> getFileFromExcel(WorksheetBase sheet1)
        {
            Dictionary<string, int> fileDict = new Dictionary<string, int>();
            int i = 8;
            bool flag = false;
            do
            {
                var range = sheet1.get_Range("A" + i.ToString());

                string stringArray = range.Cells.Value;
                if (range.Count != 0 && 
                    !string.IsNullOrWhiteSpace(stringArray))
                {
                    flag = true;
                    fileDict[stringArray] = i;
                }
                else
                {
                    flag = false;
                    break;
                }
                    
                i++;
               

            } while (flag);
            return fileDict;
        }

        internal static string[] ConvertToStringArray(System.Array values)
        {
            string[] theArray = new string[values.Length];
            for (int i = 1; i <= values.Length; i++)
            {
                if (values.GetValue(1, i) == null)
                    theArray[i - 1] = "";
                else
                    theArray[i - 1] = (string)values.GetValue(1, i).ToString();
            }
            return theArray;
        }

        public static bool UpdateDirectory(WorksheetBase sheet1)
        {
            string rootDir = sheet1.Cells[3,1].Value.ToString();
            if (string.IsNullOrWhiteSpace(rootDir))
            {
                return false;
            }

            Dictionary<string, List<string>> filesFromDir = GetFiles(rootDir);
            Dictionary<string, int> oldFilesInExcel = getFileFromExcel(sheet1);
            int cellLabel = oldFilesInExcel.LastOrDefault().Value;

            foreach(KeyValuePair<string, List<string>> dictitem in filesFromDir)
            {
                if (oldFilesInExcel.ContainsKey(dictitem.Key))
                {
                    int indexLabel = oldFilesInExcel[dictitem.Key];
                    var range = sheet1.get_Range(string.Format("B{0}:C{0}", indexLabel.ToString()));
                    range.set_Value(Missing.Value, new String[] { dictitem.Value[1], dictitem.Value[2]});
                    // update the value in excel
                    continue;
                }
                // insert a new value into 
                {
                    cellLabel += 1;
                    var range = sheet1.get_Range(string.Format("A{0}:C{0}", cellLabel.ToString()));
                    range.set_Value(Missing.Value, dictitem.Value.Take(3).ToArray());
                    InsertLinkValue(sheet1, string.Format("D{0}", cellLabel), dictitem.Value[3]);
                }
            }
            return true;
        }
    }
}
