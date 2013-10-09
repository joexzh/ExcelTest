using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using NPOI.SS.UserModel;

namespace ExcelTest
{
    public static class ExcelExp
    {
        static HSSFWorkbook hssfWorkbook;

        public static void InitiallizedWrokBook(string path)
        {
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfWorkbook = new HSSFWorkbook(fs);
            }
        }

        public static void ConvertToDataTable()
        {
            var sheet = hssfWorkbook.GetSheetAt(0);
            var rows = sheet.GetRowEnumerator();
            DataTable dt = new DataTable();
            for (int j = 0; j < 5; j++)
            {
                dt.Columns.Add(Convert.ToChar((int)'A' + j).ToString());
            }
            while (rows.MoveNext())
            {
                HSSFRow row = (HSSFRow)rows.Current;
                DataRow dr = dt.NewRow();
                for (int i = 0; i < row.LastCellNum; i++)
                {
                    var cell = row.GetCell(i);
                    if (cell == null)
                    {
                        dr[i] = null;
                    }
                    else
                    {
                        dr[i] = cell.ToString();
                    }
                }
                dt.Rows.Add(dr);
            }
        }
    }
}
