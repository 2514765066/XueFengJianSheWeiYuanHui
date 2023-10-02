namespace 学风建设委员会表格脚本
{
    enum Type
    {
        Help,
        Operate,
        Tip,
    }
    internal class Cs
    {
        //打印
        public static void Log(Type type, string msg)
        {
            switch (type)
            {
                case Type.Help:
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.Write("帮助>");
                    break;
                case Type.Operate:
                    Console.ForegroundColor = ConsoleColor.Blue;
                    Console.Write("操作>");
                    break;
                case Type.Tip:
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.Write("提示>");
                    break;
            }
            Console.WriteLine(msg);
            Console.ForegroundColor = ConsoleColor.White;
        }

        //打印提示输入
        public static string Tip(bool isEmpty, string msg)
        {
            string info;
            do
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.Write($"{msg}>");
                Console.ForegroundColor = ConsoleColor.White;
                info = Console.ReadLine()!;
            } while (isEmpty == false && info == "");            
            return info;
        }

        /// <summary>
        /// 字母变数字
        /// </summary>
        /// <param name="word">字母A=1,B=2</param>
        /// <returns></returns>
        public static int WordToNum(char word)
        {
            return Convert.ToInt32(word) - 64;
        }

        /// <summary>
        /// 数字变字母
        /// </summary>
        /// <param name="num">数字1=A,2=B</param>
        /// <returns></returns>
        public static char NumToWord(int num)
        {
            return Convert.ToChar(64 + num);
        }

        /// <summary>
        /// 字母下一个字母
        /// </summary>
        /// <param name="word">字母</param>
        /// <returns></returns>
        public static char WordNextWord(char word)
        {
            int num = WordToNum(word);
            return NumToWord(num + 1);
        }
    }
}
