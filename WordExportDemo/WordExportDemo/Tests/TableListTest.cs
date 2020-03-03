using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace WordExportDemo.Tests
{
    public class TableListTest
    {
        public static void Test(int count = 10)
        {
            Console.WriteLine($"表格");
            Stopwatch sw = new Stopwatch();
            IList<Dictionary<string, object>> data = new List<Dictionary<string, object>>();
            for (int i = 0; i < count; i++)
            {
                Dictionary<string, object> d = new Dictionary<string, object>()
                {
                    { "xm","测试"+i.ToString()},
                    { "nl",i},
                    { "xb","男"}
                };
                data.Add(d);
            }

            Dictionary<string, object> data1 = new Dictionary<string, object>();
            data1.Add("list", data);
            sw.Start();
            string root = System.AppDomain.CurrentDomain.BaseDirectory;
            WordHelper.Export(root + Path.Combine("Templates", "temp_table_list.docx"), root + "temp_table_list_out.docx", data1);
            sw.Stop();
            var time = sw.ElapsedMilliseconds;
            Console.WriteLine($"耗时：{time}毫秒");
        }
    }
}