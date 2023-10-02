using System.Diagnostics;
using OfficeOpenXml;

namespace 学风建设委员会表格脚本
{
    internal class Program
    {
        const string v = "3.1.1";
        static void Main()
        {
            //epplus免费许可证
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            Console.WriteLine($"欢迎使用学风建设委员会表格制作脚本\n当前版本:{v}\n");
            do
            {
                Console.Write("用户>");
                //输入
                string command = Console.ReadLine()!;
                //名字
                string name;

                switch (command)
                {
                    case "help":
                        Cs.Log(Type.Help, "clear(清空控制台)\n");
                        Cs.Log(Type.Help, "open(打开核算表模板)\n");
                        Cs.Log(Type.Help, "合并表格(合并源文件夹中所有表格)\n");
                        Cs.Log(Type.Help, "制作未打卡日表\n");
                        Cs.Log(Type.Help, "制作未打卡周表\n");
                        Cs.Log(Type.Help, "制作核算表\n");
                        Cs.Log(Type.Help, "制作扣分表\n");
                        Cs.Log(Type.Help, "删除名单\n");
                        break;
                    case "clear":
                        Console.Clear();
                        break;
                    case "open":
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = "核算表模板.xlsx",
                            UseShellExecute = true,
                        });

                        Cs.Log(Type.Operate, "已打开核算表模板\n");
                        break;
                    case "合并表格":
                        //输入的名字
                        name = Cs.Tip(true, "输入合并后的文件名(按下enter使用默认名字)");

                        Operate.MergeExcel(name == "" ? "合并后的表格" : name);

                        Cs.Log(Type.Operate, "合并表格完成\n");
                        break;
                    case "制作未打卡日表":
                        Make.MakeDayExcel();

                        Cs.Log(Type.Operate, "未打卡日表制作完成\n");
                        break;
                    case "制作未打卡周表":
                        Make.MakeWeekExcel();

                        Cs.Log(Type.Operate, "未打卡周表制作完成\n");
                        break;
                    case "制作核算表":
                        Make.Func accounting = new(Make.MakeAccountingExcel);
                        Make.MakeDeductTemplate(accounting);

                        Cs.Log(Type.Operate, "核算表制作完成\n");
                        break;
                    case "制作扣分表":
                        Make.Func deduct = new(Make.MakeDeductExcel);
                        Make.MakeDeductTemplate(deduct);

                        Cs.Log(Type.Operate, "扣分表制作完成\n");
                        break;
                    case "删除名单":
                        Make.DeleteName();

                        Cs.Log(Type.Operate, "名单已删除\n");
                        break;
                    default:
                        Cs.Log(Type.Tip, "未知命令,输入help查看全部命令\n");
                        break;
                }
            } while (true);
        }
    }
}