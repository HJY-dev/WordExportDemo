using System;

namespace WordExportDemo
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            Tests.FormTableTest.Test();
            Tests.FormTest.Test();
            Tests.TableListTest.Test();
            Tests.WtTest.Test();

            Console.WriteLine("测试完成");
            Console.ReadKey();
        }
    }
}