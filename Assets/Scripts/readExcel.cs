using System.Collections.Generic;
using System.Data;
using UnityEngine;
using System.IO;
using Excel;


public class readExcel
{
    public static void readexcel()
    {
        FileStream fileStream =
            File.Open(Application.streamingAssetsPath + "/writeTemplate", FileMode.Open, FileAccess.Read);

        IExcelDataReader excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(fileStream);
        DataSet result = excelDataReader.AsDataSet();

        // 获取表格有多少列
        int columns = result.Tables[0].Columns.Count;

        // 获取表格有多少行 
        int rows = result.Tables[0].Rows.Count;

        // 根据行列依次打印表格中的每个数据 
        List<string> excelDta = new List<string>();

        //第一行为表头，不读取。没有表头从0开始
        for (int i = 1; i < rows; i++)
        {
            string value = null;
            string all = null;
            for (int j = 0; j < columns; j++)
            {
                // 获取表格中指定行指定列的数据 
                value = result.Tables[0].Rows[i][j].ToString();
                if (value == "")
                {
                    continue;
                }

                all = all + value + "|";
            }

            if (all != null)
            {
                //print(all);
                excelDta.Add(all);
            }
        }
        string[] item = data_1[i].Split('|');
        PlayerInfo playerinfo = new PlayerInfo(item[0], item[1], item[2], mobile_Enc);
    }
    
}