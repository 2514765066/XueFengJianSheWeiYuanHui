using OfficeOpenXml;
using System.Data;

namespace 学风建设委员会表格脚本
{
    internal class Operate
    {
        //合并表格
        public static string MergeExcel(string newExcelName)
        {
            string dir_path = Cs.Tip(false, "输入需要合并的文件夹路径").Replace("\"", null);
            //创建一个新的excel表格
            using ExcelPackage newExcel = new();

            //添加一个sheet1
            ExcelWorksheet sheet = newExcel.Workbook.Worksheets.Add("sheet1");

            //遍历文件夹中所有表格
            foreach (var file in Directory.GetFiles(dir_path))
            {
                //打开源表格
                using ExcelPackage sourceExcel = new(new FileInfo(file));
                //选择第一个sheet
                ExcelWorksheet sourceSheet = sourceExcel.Workbook.Worksheets[0];
                //获取范围
                ExcelRange sourceRange = sourceSheet.Cells[sourceSheet.Dimension.Address];
                //最后一行
                int lastRow = sheet.Dimension?.End.Row ?? 0;

                ExcelRange destRange = sheet.Cells[lastRow + 1, 1, lastRow + sourceRange.Rows, sourceRange.Columns];
                sourceRange.Copy(destRange);
            }

            string newExcel_path = dir_path + @$"\{newExcelName}.xlsx";
            newExcel.SaveAs(new FileInfo(newExcel_path));
            return newExcel_path;
        }

        //替换所有单元格中特定的内容
        public static void ClearChar(string path, string[] clearChar)
        {
            //打开源表格
            using ExcelPackage sourceExcel = new(new FileInfo(path));
            //选择第一个sheet
            ExcelWorksheet sourceSheet = sourceExcel.Workbook.Worksheets[0];
            //获取范围
            ExcelRange sourceRange = sourceSheet.Cells[sourceSheet.Dimension.Address];

            foreach (var cell in sourceRange)
            {

                // 将新值赋给单元格                
                foreach (string item in clearChar)
                {
                    // 获取单元格的值
                    string value = cell.Value.ToString()!;
                    cell.Value = value.Replace(item, null);
                }
            }
            sourceExcel.Save();
        }

        /// <summary>
        /// 删除行
        /// </summary>
        /// <param name="path">路径</param>
        /// <param name="name">列名</param>
        public static void DeleteColumn(string path, char[] columnName)
        {
            int[] index = columnName.Select((c, i) => c - 64 - i).ToArray();

            //打开源表格
            using ExcelPackage sourceExcel = new(new FileInfo(path));
            //选择第一个sheet
            ExcelWorksheet sourceSheet = sourceExcel.Workbook.Worksheets[0];

            foreach (int item in index)
            {
                sourceSheet.DeleteColumn(item);
            }
            sourceExcel.Save();
        }

        /// <summary>
        /// 删除行
        /// </summary>
        /// <param name="sourceExcel">表格</param>
        /// <param name="columnName">列名</param>
        public static void DeleteColumn(ExcelWorksheet sourceSheet, char[] columnName)
        {
            int[] index = columnName.Select((c, i) => c - 64 - i).ToArray();

            foreach (int item in index)
            {
                sourceSheet.DeleteColumn(item);
            }
        }

        /// <summary>
        /// 剪切列
        /// </summary>
        /// <param name="path">路径</param>
        /// <param name="from">来自哪一列</param>
        /// <param name="to">去自哪一列</param>
        public static void CutColumn(string path, char from, char to)
        {
            //打开源表格
            using ExcelPackage sourceExcel = new(new FileInfo(path));
            //选择第一个sheet
            ExcelWorksheet sourceSheet = sourceExcel.Workbook.Worksheets[0];
            //插入一列
            sourceSheet.InsertColumn(Cs.WordToNum(to), 1);
            //选区从前往后
            ExcelRange sourceRange = sourceSheet.Cells[$"{from}:{from}"];
            //从后往前剪切
            if (Cs.WordToNum(from) > Cs.WordToNum(to))
            {
                sourceRange = sourceSheet.Cells[$"{Cs.WordNextWord(from)}:{Cs.WordNextWord(from)}"];
            }
            sourceRange.Copy(sourceSheet.Cells[$"{to}:{to}"]);
            sourceRange.Delete(eShiftTypeDelete.Left);
            sourceExcel.Save();
        }

        //排序
        public static void Sort(string path)
        {
            //打开源表格
            using ExcelPackage sourceExcel = new(new FileInfo(path));
            //选择第一个sheet
            ExcelWorksheet sourceSheet = sourceExcel.Workbook.Worksheets[0];
            sourceSheet.Cells[sourceSheet.Dimension.Address].Offset(1, 0).Sort(new int[] { 0, 1 }, new bool[] { false, false });
            sourceExcel.Save();
        }

        //删除重复值
        public static void ClearReset(string path)
        {
            //打开源表格
            using ExcelPackage sourceExcel = new(new FileInfo(path));
            //选择第一个sheet
            ExcelWorksheet sourceSheet = sourceExcel.Workbook.Worksheets[0];

            DataTable dataTable = sourceSheet.Cells[sourceSheet.Dimension.Address].ToDataTable(options => { options.FirstRowIsColumnNames = false; }).AsEnumerable().Distinct(DataRowComparer.Default).CopyToDataTable();
            sourceSheet.Cells.Clear();
            sourceSheet.Cells["A1"].LoadFromDataTable(dataTable, true);
            sourceSheet.DeleteRow(1);
            sourceExcel.Save();
        }

        //单元格自适应大小
        public static void AutoSize(string path)
        {
            //打开源表格
            using ExcelPackage sourceExcel = new(new FileInfo(path));
            //选择第一个sheet
            ExcelWorksheet sourceSheet = sourceExcel.Workbook.Worksheets[0];

            sourceSheet.Cells[sourceSheet.Dimension.Address].AutoFitColumns();
            sourceExcel.Save();
        }

        //单元格自适应大小
        public static void AutoSize(ExcelWorksheet sourceSheet)
        {
            sourceSheet.Cells[sourceSheet.Dimension.Address].AutoFitColumns();
        }
    }
}
