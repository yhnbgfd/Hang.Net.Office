using Hang.Net.Office.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                Console.WriteLine("========================");
                Console.WriteLine("i：word文档转化成图片");
                Console.WriteLine("x：关闭程序");
                Console.Write("请输入功能码：");
                var key = Console.ReadKey().KeyChar;
                Console.WriteLine();
                TestOffice(key);
            }
        }

        static void TestOffice(char key)
        {
            Console.WriteLine($">>>>>> Test office : {key}");
            switch (key)
            {
                case 'i':
                    MsWordUtility.ToImage(@"D:\北京平安力合科技发展股份有限公司\排队机\0-平安力合智能排队管理系统(SOMS520&CQ510R5)对外接口通讯协议说明（新后台修改）.doc", @"C:\Users\yhnbg\Desktop\a");
                    break;
                case 'x':
                    Environment.Exit(0);
                    break;
                default:
                    break;
            }
        }
    }
}
