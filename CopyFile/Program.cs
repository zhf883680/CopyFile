using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace CopyFile
{
    class Program
    {
        static void Main(string[] args)
        {
            //设置文件夹路径
            string originPath = @"I:\Music";//源文件路径
            string targetPath = @"E:\音乐\iTunes\经典\";//目标文件路径
            //获取文件名
            var dt = ExcelToDataTable("music", true, @"E:\桌面\music.xlsx");//导出的excel  此处为itunes导出的txt 后用excel打开另存为格式
            List<string> filenames = new List<string>();
            foreach (DataRow dr in dt.Rows)
            {
                filenames.Add(dr[0] as string);
            }
            Console.WriteLine("已获取歌名列表");
            //遍历源文件夹中所有音乐
            DirectoryInfo theFolder = new DirectoryInfo(originPath);
            var files = theFolder.GetFiles();
            Console.WriteLine("已获取文件夹中文件列表");
            var filename = string.Empty;
            Console.WriteLine("遍历文件夹中文件");
            foreach (var file in files)
            {
                filename = file.Name.Substring(file.Name.IndexOf(' ')+3);//获取音乐名
                filename = filename.Substring(0, filename.Length - 4);//此处依网易云格式 周杰伦 - 床边故事.MP3
                foreach (var name in filenames)
                {
                    if (filename == name)
                    {
                        Console.WriteLine("找到同名文件,正在复制" + file.Name);
                        try
                        {
                            file.CopyTo(targetPath + file.Name);
                        }
                        catch(Exception ex)
                        {
                            Console.WriteLine(file.Name+ex.Message);
                        }
                        
                        continue;
                    }
                }
            }
            Console.WriteLine("文件复制结束,按任意键关闭窗口");
            Console.ReadLine();
        }
        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="sheetName">excel工作薄sheet的名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>返回的DataTable</returns>
        static DataTable ExcelToDataTable(string sheetName, bool isFirstRowColumn, string fileName)
        {
            IWorkbook workbook = null;
            FileStream fs = null;
            ISheet sheet = null;
            DataTable data = new DataTable();
            int startRow = 0;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                    workbook = new XSSFWorkbook(fs);
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                    workbook = new HSSFWorkbook(fs);

                if (sheetName != null)
                {
                    sheet = workbook.GetSheet(sheetName);
                    if (sheet == null) //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                    {
                        sheet = workbook.GetSheetAt(0);
                    }
                }
                else
                {
                    sheet = workbook.GetSheetAt(0);
                }
                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数

                    if (isFirstRowColumn)
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            ICell cell = firstRow.GetCell(i);
                            if (cell != null)
                            {
                                string cellValue = cell.StringCellValue;
                                if (cellValue != null)
                                {
                                    DataColumn column = new DataColumn(cellValue);
                                    data.Columns.Add(column);
                                }
                            }
                        }
                        startRow = sheet.FirstRowNum + 1;
                    }
                    else
                    {
                        startRow = sheet.FirstRowNum;
                    }

                    //最后一列的标号
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null　　　　　　　

                        DataRow dataRow = data.NewRow();
                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                        {
                            if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                                dataRow[j] = row.GetCell(j).ToString();
                        }
                        data.Rows.Add(dataRow);
                    }
                }

                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return null;
            }
        }

    }
}
