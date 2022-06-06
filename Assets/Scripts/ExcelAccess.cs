// using UnityEngine;
// using Excel;
// using System.Data;
// using System.IO;
// using OfficeOpenXml;
//
// public class ExcelAccess
// {
//
//
//     #region 读
//     //读取的路径上的文件夹
//     static string readFile= "C:/moban";
//     //读取路径上的文件xlsx
//     static string readPath = readFile+"/moban1.xlsx";
//     public static string[,] ReadExcel(out byte n)
//     {
//         if (!Directory.Exists(readFile))
//         {
//             Directory.CreateDirectory(readFile);
//         }
//         if (!File.Exists(readPath))
//         {
//             n = 1;
//             return null;//没有目标文件！
//         }
//         //FileShare.ReadWrite 可以打开没有关闭的excel
//         FileStream fs = File.Open(readPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//         IExcelDataReader excelRead = ExcelReaderFactory.CreateOpenXmlReader(fs);
//         DataSet result = excelRead.AsDataSet();
//         //行 Tables[0]表示的是第一页！
//         int row = result.Tables[0].Rows.Count;
//         //列
//         int columns = result.Tables[0].Columns.Count;
//         //输出
//         string[,] resultstrarray = new string[row, columns];
//         for (int i = 0; i < row; i++)
//         {
//             for (int j = 0; j < columns; j++)
//             {
//                 string str = result.Tables[0].Rows[i].ItemArray[j].ToString();
//                 //去空格
//                 str = str.Replace(" ", "");
//                 //加入数组
//                 resultstrarray[i, j] = str;
//                 //Debug.Log("i: " + i + "J: " + j + "str: " + str);
//             }
//         }
//     
//         n = 0;
//         return resultstrarray;//成功读取
//
//     }
//
//     #endregion
//
//     #region 写
//
//     static string writeFile= "C:/moban";
//     static string writePath = writeFile + "/writemoban.xlsx";
//     //直接写
//     public static void WriteExcel()
//     {
//         if (!Directory.Exists(writeFile))
//         {
//             Directory.CreateDirectory(writeFile);
//         }
//         if (File.Exists(writePath))
//         {
//             Debug.Log("文件已经存在！！删除");
//             File.Delete(writePath);
//         }
//         FileInfo newFile = new FileInfo(writePath);
//         using (ExcelPackage package = new ExcelPackage(newFile))
//         {
//             ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("mk");
//             //直接写
//             worksheet.Cells[1, 1].Value = "ID";
//
//             worksheet.Cells["D3"].Value = 999;
//             //合并单元格 1,1  和1 6 合并
//             worksheet.Cells[1, 1, 1, 6].Merge = true;
//             //设置样式大小和颜色
//             worksheet.Cells["F10"].Value = "看这里";
//             worksheet.Cells["F10"].Style.Font.Size = 20;
//             worksheet.Cells["F10"].Style.Font.Name = "宋体";
//             worksheet.Cells["F10"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
//             //保存excel
//             package.Save();
//         }
//
//     }
//
//     public static void WriteExcelTemplate()
//     {
//         string sourcePath = Application.streamingAssetsPath + "/writeTemplate.xlsx";
//         string targetPath= writeFile + "/writeTemplate.xlsx";
//
//         if (File.Exists(targetPath))
//         {
//             Debug.Log("文件已经存在！！删除");
//             File.Delete(targetPath);
//         }
//         using (ExcelPackage sourcePack=new ExcelPackage(new FileInfo(sourcePath)))
//         {
//             using (ExcelPackage targetPack = new ExcelPackage(new FileInfo(targetPath)))
//             {
//                 //注意sourcePack.Workbook.Worksheets[1]这个是1
//                 ExcelWorksheet worksheet = targetPack.Workbook.Worksheets.Add("s", sourcePack.Workbook.Worksheets[1]);
//                 worksheet.Cells[2, 1].Value = "12223";
//                 worksheet.Cells[2, 2].Value = "白雪";
//                 worksheet.Cells[2, 3].Value = "100";
//                 targetPack.Save();
//             }
//         }
//     }
//     #endregion
//     /// <summary>
//     /// 写入 Excel ; 需要添加 OfficeOpenXml.dll;
//     /// </summary>
//     /// <param name="excelName">excel文件名</param>
//     /// <param name="sheetName">sheet名称</param>
//     public static void WriteExcel(string excelName, string sheetName)
//     {
//         //通过面板设置excel路径
//         //string outputDir = EditorUtility.SaveFilePanel("Save Excel", "", "New Resource", "xlsx");
//
//         //自定义excel的路径
//         string path = Application.dataPath + "/" + excelName;
//         FileInfo newFile = new FileInfo(path);
//         if (newFile.Exists)
//         {
//             //创建一个新的excel文件
//             newFile.Delete();
//             newFile = new FileInfo(path);
//         }
//
//         //通过ExcelPackage打开文件
//         using (ExcelPackage package = new ExcelPackage(newFile))
//         {
//             //在excel空文件添加新sheet
//             ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);
//             //添加列名
//             worksheet.Cells[1, 1].Value = "ID";
//             worksheet.Cells[1, 2].Value = "Product";
//             worksheet.Cells[1, 3].Value = "Quantity";
//             worksheet.Cells[1, 4].Value = "Price";
//             worksheet.Cells[1, 5].Value = "Value";
//
//             //添加一行数据
//             worksheet.Cells["A2"].Value = 12001;
//             worksheet.Cells["B2"].Value = "Nails";
//             worksheet.Cells["C2"].Value = 37;
//             worksheet.Cells["D2"].Value = 3.99;
//             //添加一行数据
//             worksheet.Cells["A3"].Value = 12002;
//             worksheet.Cells["B3"].Value = "Hammer";
//             worksheet.Cells["C3"].Value = 5;
//             worksheet.Cells["D3"].Value = 12.10;
//             //添加一行数据
//             worksheet.Cells["A4"].Value = 12003;
//             worksheet.Cells["B4"].Value = "Saw";
//             worksheet.Cells["C4"].Value = 12;
//             worksheet.Cells["D4"].Value = 15.37;
//
//             //保存excel
//             package.Save();
//         }
//     }
//
// }
//
