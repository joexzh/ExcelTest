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
        static DataTable dt = new DataTable();

        public static void InitiallizedWrokBook(string path)
        {
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfWorkbook = new HSSFWorkbook(fs);
            }
        }

        //public static void Convet

        public static void ConvertToDataTable()
        {
            var sheet = hssfWorkbook.GetSheetAt(0);
            var rows = sheet.GetRowEnumerator();
            dt = new DataTable();
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

        public static void ExpToXls()
        {
            var sheet1 = hssfWorkbook.GetSheetAt(0);
            sheet1.GetRow(1).GetCell(1).SetCellValue(200020);
            sheet1.GetRow(2).GetCell(1).SetCellValue(200021);
            sheet1.GetRow(3).GetCell(1).SetCellValue(200022);
            sheet1.GetRow(4).GetCell(1).SetCellValue(200023);
            sheet1.GetRow(5).GetCell(1).SetCellValue(200024);
            sheet1.GetRow(6).GetCell(1).SetCellValue(200025);
            sheet1.GetRow(7).GetCell(1).SetCellValue(200026);
            sheet1.GetRow(8).GetCell(1).SetCellValue(200027);
            sheet1.GetRow(9).GetCell(1).SetCellValue(200028);
            sheet1.GetRow(10).GetCell(1).SetCellValue(200029);
            sheet1.GetRow(11).GetCell(1).SetCellValue(200030);
            sheet1.GetRow(12).GetCell(1).SetCellValue(200031);

            sheet1.ForceFormulaRecalculation = true;

            FileStream fs = new FileStream(@"e:\ExpToXls.xls", FileMode.Create);
            hssfWorkbook.Write(fs);
            fs.Close();
        }

        public static void CreateSheetAndExp(DataTable dt)
        {
            hssfWorkbook = new HSSFWorkbook();
            var sheet = hssfWorkbook.CreateSheet("sheet1");
            var cellStyle = hssfWorkbook.CreateCellStyle();
            cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");

            var stringStyle = hssfWorkbook.CreateCellStyle();
            stringStyle.VerticalAlignment = VerticalAlignment.CENTER;

            int columnCount = dt.Columns.Count;
            int[] arrColWidth = new int[columnCount];
            int width = 10;

            foreach (DataColumn item in dt.Columns)
            {
                arrColWidth[item.Ordinal] = width;
            }


        }
    }
}
