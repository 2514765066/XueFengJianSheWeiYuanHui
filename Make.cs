using OfficeOpenXml;

namespace 学风建设委员会表格脚本
{
    internal class Make
    {
        public delegate void Func(ExcelPackage sourceExcel, string excel_path);

        //制作日表
        public static void MakeDayExcel()
        {
            DateTime today = DateTime.Today;

            string excel_path = Operate.MergeExcel($"{today.Month}月{today.Day}日未打卡名单");

            Operate.ClearChar(excel_path, new string[] { "\t", "定位签到", "精准定位" });
            Operate.DeleteColumn(excel_path, new char[] { 'C', 'D', 'L' });
            Operate.CutColumn(excel_path, 'C', 'A');
            Operate.Sort(excel_path);
            Operate.ClearReset(excel_path);
            Operate.AutoSize(excel_path);
        }

        //制作周表
        public static void MakeWeekExcel()
        {
            string excel_path = Operate.MergeExcel("打卡第x周");

            Operate.ClearChar(excel_path, new string[] { "\t", "定位签到", "精准定位" });
            Operate.ClearReset(excel_path);
            Operate.Sort(excel_path);
            Operate.AutoSize(excel_path);
        }

        //制作扣分表模板
        public static void MakeDeductTemplate(Func func)
        {
            string excel_path = Cs.Tip(false, "输入未打卡周表的路径").Replace("\"", null);
            Operate.Sort(excel_path);
            //打开源表格
            using ExcelPackage sourceExcel = new(new FileInfo(@excel_path));
            //选择第一个sheet
            ExcelWorksheet sourceSheet = sourceExcel.Workbook.Worksheets[0];
            //最后一行的索引
            int lastRow = sourceSheet.Dimension.End.Row;

            //统计异常次数
            sourceSheet.Cells["J1"].Value = "异常次数";
            for (int i = 2; i <= lastRow; i++)
            {
                sourceSheet.Cells[$"J{i}"].Formula = $"COUNTIF(D{i}:H{i},\"异常\")";
                sourceSheet.Calculate();
                sourceSheet.Cells[$"J{i}"].Value = sourceSheet.Cells[$"J{i}"].Text;
            }


            //统计人名*异常总次数
            sourceSheet.Cells["K1"].Value = "人名*异常总次数";
            int sum = 0;
            for (int i = 2; i <= lastRow; i++)
            {
                // 获取当前行、上一行和下一行的C列单元格的文本值
                string current = sourceSheet.Cells[$"C{i}"].Text;
                string previous = sourceSheet.Cells[$"C{i - 1}"].Text;
                string next = sourceSheet.Cells[$"C{i + 1}"].Text;

                // 获取当前行的B列和J列单元格的值
                object bValue = sourceSheet.Cells[$"B{i}"].Value;
                object jValue = sourceSheet.Cells[$"J{i}"].Value;

                //上不等下不等
                if (current != previous && current != next)
                {
                    sourceSheet.Cells[$"K{i}"].Value = $"{bValue}*{jValue}";
                }
                //上不等下等和上等下等
                else if ((current != previous && current == next) || (current == previous && current == next))
                {
                    sum += Convert.ToInt32(jValue);
                }
                //上等下不等
                else
                {
                    sourceSheet.Cells[$"K{i}"].Value = $"{bValue}*{sum + Convert.ToInt32(jValue)}";
                    sum = 0;
                }
            }


            //班级总扣分和名单
            sourceSheet.Cells[$"L1"].Value = "扣分";
            sourceSheet.Cells[$"M1"].Value = "班级扣分名单";
            //整个班级的起始行
            int startrow = 2;
            for (int i = 2; i <= lastRow; i++)
            {
                string current = sourceSheet.Cells[$"A{i}"].Text;
                string previous = sourceSheet.Cells[$"A{i - 1}"].Text;
                string next = sourceSheet.Cells[$"A{i + 1}"].Text;

                //与上相等与下不相等
                if (current == previous && current != next)
                {
                    int all = 0;
                    //泛型数组存储有数据的单元格
                    List<string> rows = new();
                    for (int j = startrow; j <= i; j++)
                    {
                        all += Convert.ToInt32(sourceSheet.Cells[$"J{j}"].Value);
                        rows.Add($"K{j}");
                    }
                    //在班级最后的位置写入扣分和未打卡名单*次数
                    sourceSheet.Cells[$"M{i}"].Formula = string.Join('&', rows);
                    sourceSheet.Calculate();
                    sourceSheet.Cells[$"M{i}"].Value = sourceSheet.Cells[$"M{i}"].Value.ToString()!;

                    sourceSheet.Cells[$"L{i}"].Value = all.ToString();
                    startrow = i + 1;
                }
                //上不等下不等只有一个人的情况
                else if (current != previous && current != next)
                {
                    sourceSheet.Cells[$"M{i}"].Value = sourceSheet.Cells[$"K{i}"].Value;
                    sourceSheet.Cells[$"L{i}"].Value = sourceSheet.Cells[$"J{i}"].Value.ToString();
                    startrow++;
                }
            }

            func(sourceExcel, excel_path);
        }

        //制作核算表
        public static void MakeAccountingExcel(ExcelPackage sourceExcel, string excel_path)
        {
            ExcelWorksheet sourceSheet = sourceExcel.Workbook.Worksheets[0];

            using ExcelPackage templateExcel = new(new FileInfo("核算表模板.xlsx"));
            ExcelWorksheets templateSheets = templateExcel.Workbook.Worksheets;

            //把所有扣分填入
            foreach (var cell in sourceSheet.Cells["L:L"])
            {
                if (cell.Text == "扣分") continue;

                string class_name = sourceSheet.Cells[cell.Start.Row, 1].Text;
                object dedcut_point = sourceSheet.Cells[cell.Start.Row, 12].Value;
                object all_name = sourceSheet.Cells[cell.Start.Row, 13].Value;

                foreach (var sheet in templateSheets)
                {
                    foreach (var Acell in sheet.Cells["A:A"])
                    {
                        if (Acell.Text != class_name) continue;

                        sheet.Cells[Acell.Start.Row, 4].Value = dedcut_point;
                        sheet.Cells[Acell.Start.Row, 2].Value = all_name;
                        break;
                    }
                }
            }

            string newExcel_path = Path.GetDirectoryName(excel_path) + @"\学风建设委员会第x周.xlsx";
            templateExcel.SaveAs(newExcel_path);
        }

        //制作扣分表
        public static void MakeDeductExcel(ExcelPackage sourceExcel, string excel_path)
        {
            string date = Cs.Tip(false, "输入扣分表日期");
            ExcelWorksheet sourceSheet = sourceExcel.Workbook.Worksheets[0];
            int lastRow = sourceSheet.Dimension.End.Row;

            sourceSheet.Cells["D1"].Value = "原因";
            sourceSheet.Cells["E1"].Value = "分值";
            for (int i = 2; i <= lastRow; i++)
            {
                if (sourceSheet.Cells[$"K{i}"].Value == null)
                {
                    sourceSheet.DeleteRow(i);
                    i--;
                    lastRow--;
                    continue;
                }
                string count = sourceSheet.Cells[$"K{i}"].Text.Split("*")[1];
                sourceSheet.Cells[$"D{i}"].Value = $"{date}未打卡{count}次";
                sourceSheet.Cells[$"E{i}"].Value = $"-{int.Parse(count) * 0.2m}";
            }
            Operate.DeleteColumn(sourceSheet, new char[] { 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M' });
            Operate.AutoSize(sourceSheet);

            string newExcel_path = Path.GetDirectoryName(excel_path) + @"\扣分表第x周.xlsx";
            sourceExcel.SaveAs(newExcel_path);
        }

        //删除名单
        public static void DeleteName()
        {
            string excel_path = Cs.Tip(false, "输入表格的路径").Replace("\"", null);
            string delete_path = Cs.Tip(false, "输入需要删除的名单的路径").Replace("\"", null);


            using ExcelPackage sourceExcel = new(new FileInfo(excel_path));
            ExcelWorksheet sourceSheet = sourceExcel.Workbook.Worksheets[0];

            using ExcelPackage deleteExcel = new(new FileInfo(delete_path));
            ExcelWorksheet deleteSheet = deleteExcel.Workbook.Worksheets[0];

            ExcelRange excelRange = sourceSheet.Cells["C:C"];


            for (int i = 1; i <= sourceSheet.Dimension.End.Row; i++)
            {
                foreach (var xuehao in deleteSheet.Cells["C:C"])
                {
                    if (sourceSheet.Cells[$"C{i}"].Text == "学号") continue;
                    //if (sourceSheet.Cells[$"C{i}"].Text != xuehao.Text) continue;
                    if ( xuehao.Text.IndexOf(sourceSheet.Cells[$"C{i}"].Text) == -1) continue;

                    sourceSheet.DeleteRow(i);
                    i--;
                }
            }
            sourceExcel.Save();
        }
    }
}
