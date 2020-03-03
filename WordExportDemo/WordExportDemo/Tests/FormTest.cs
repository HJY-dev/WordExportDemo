using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace WordExportDemo.Tests
{
    public class FormTest
    {
        public static void Test()
        {
            Console.WriteLine($"表单");
            Stopwatch sw = new Stopwatch();
            Dictionary<string, object> data = new Dictionary<string, object>()
            {
                { "xmmc","测试姓名测试姓名111"},
                { "sqje","1417.4"},
                { "xmdw","博客园Deeround"},
                { "glfs","自行管理方式"},
                { "xmgk","博客园Deeround来函申请办理 应急抢险治理工程项目竣工结（决）算，该项目已完工并通过项目初步验收，现拟按程序采取政府购买服务方式开展评审"},
                { "psyj",""},
                { "gzyq", @"（一）对建设程序进行评审，包括可行性研究报告、初步设计等批准文件的程序性审查。
（二）对建设规模、建设标准、可研执行情况等进行评审。
（三）对工程投资进行评审，包括工程计量、定额选用、材料价格及费用标准等的评审。
（四）对设施设备资进行评审，包括设施设备型号、规格、数量及价格的评审。
"},
                { "wcsx","1. 收到委托书后在10天内报送评审方案，评审完成后需提交评审报告纸质件7份及电子文档。"},
                { "ywcs","伯爵二元"},
                { "lxr","千年  12345678"},
            };

            sw.Start();
            string root = System.AppDomain.CurrentDomain.BaseDirectory;
            WordHelper.Export(root + Path.Combine("Templates", "temp_form.docx"), root + "temp_form_out.docx", data);
            sw.Stop();
            var time = sw.ElapsedMilliseconds;
            Console.WriteLine($"耗时：{time}毫秒");
        }
    }
}