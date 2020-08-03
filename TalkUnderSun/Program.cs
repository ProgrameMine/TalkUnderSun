using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using EtcJob.HelpClass;
using Newtonsoft.Json;

namespace TalkUnderSun
{
    class Program
    {
        static void Main(string[] args)
        {
            var paths = new List<string>
            {
                @"H:\Ehi\Ehi.Zj.Settlement\Ehi.Zj.Settlement.Service",
                @"H:\Ehi\Ehi.ZJ.Manage\Ehi.ZJ.Manage.Service",
                @"H:\Ehi\Ehi.Zj.Violation\Ehi.Zj.Violation.Service"
            };

            var apiMothodList = new List<ModelAPIMethod>();

            foreach (var path in paths)
            {
                var pathArr = path.Split(@".", StringSplitOptions.None).ToArray();
                var systemName = pathArr[pathArr.Length - 2];

                var files = Directory.GetFiles(path, "*.cs").ToList();
                if (files == null || files.Count == 0)
                {
                    Console.WriteLine("无有效cs文件");
                    Console.ReadKey();
                }
                files = files.Where(m => !m.Contains("HttpHandler")).ToList();

                var date1 = DateTime.Now;
                Console.WriteLine("开始读取cs文件：" + systemName);
                foreach (var file in files)
                {
                    var fileArr = file.Split(@"\", StringSplitOptions.None).ToArray();
                    if (fileArr == null || fileArr.Length == 0) continue;
                    var fileName = fileArr[fileArr.Length - 1];
                    Console.WriteLine(fileName.TrimEnd('.', 'c', 's') + "正在读取...");
                    apiMothodList.AddRange(ReadWebAPIFile(file, systemName, systemName + "_" + fileName.TrimEnd('.', 'c', 's')));
                    Console.WriteLine(fileName.TrimEnd('.', 'c', 's') + "读取完毕！");
                }
                Console.WriteLine(systemName + "的cs文件读取完成！总耗时：" + (DateTime.Now - date1));
            }

            var filePath = @"C:\Users\Administrator\Desktop\新建 Microsoft Excel 工作表 (4).xlsx";
            var date2 = DateTime.Now;
            Console.WriteLine("开始读取匹配参照文件：" + filePath);
            var apiDescriptions = ReadAPIDescrption(filePath);
            Console.WriteLine("匹配参照文件读取完成！总耗时：" + (DateTime.Now - date2));

            var date3 = DateTime.Now;
            Console.WriteLine("开始匹配...");
            foreach (var item in apiDescriptions)
            {
                var apiMethods = apiMothodList.Where(m => m.APIUrl == item.API).ToList();
                if (apiMethods == null || apiMethods.Count == 0)
                {
                    item.FunctionAddr = "暂无调用";
                    continue;
                }

                if (apiMethods.Count > 0)
                {
                    item.FunctionAddr = string.Join(";", apiMethods.Select(n => n.ServiceFileName).Distinct().ToArray());
                    item.Function = string.Join(";", apiMethods.Select(n => n.Method).Distinct().ToArray());
                    item.SystemName = string.Join(";", apiMethods.Select(n => n.SystemName).Distinct().ToArray());
                }
            }
            Console.WriteLine("匹配完成！总耗时：" + (DateTime.Now - date3));

            var unUsingApiList = apiDescriptions.Where(m => m.FunctionAddr == "暂无调用").ToList();
            if (unUsingApiList != null && unUsingApiList.Count > 0)
            {
                Console.WriteLine("无调用方法数：" + unUsingApiList.Count);
            }
            Console.WriteLine(JsonConvert.SerializeObject(apiDescriptions));

            var dict = new Dictionary<string, string>()
            {
                { "API","API"},
                { "Description","Description"},
                { "Type","类型"},
                { "FunctionAddr","方法文件地址"},
                { "Function","方法名"},
                { "SystemName","系统名称"}
            };

            var dateTable = ExportTo.ConvertListToDataTable(apiDescriptions, dict);
            ExportTo.Export(dateTable, @"C:\Users\Administrator\Desktop", "");
            Console.ReadKey();
        }

        static List<ModelAPIMethod> ReadWebAPIFile(string strReadFilePath, string systemName, string fileName)
        {
            // 读取文件的源路径及其读取流
            StreamReader srReadFile = new StreamReader(strReadFilePath);

            List<ModelAPIMethod> list = new List<ModelAPIMethod>();

            var model = new ModelAPIMethod();
            // 读取流直至文件末尾结束
            while (!srReadFile.EndOfStream)
            {
                string strReadLine = srReadFile.ReadLine(); //读取每行数据
                if (strReadLine.Contains("public") && !strReadLine.Contains("class"))
                    model.Method = strReadLine.Trim();
                if (strReadLine.Contains("method ="))
                    model.APIUrl = strReadLine.Trim();
                if (!string.IsNullOrEmpty(model.Method) && !string.IsNullOrEmpty(model.APIUrl))
                {
                    list.Add(model);
                    model = new ModelAPIMethod();
                }
            }
            srReadFile.Close();

            list = list.Distinct().ToList();
            foreach (var item in list)
            {
                var methodArr = item.Method.Split("(").ToArray();
                if (methodArr != null && methodArr.Length > 1)
                {
                    item.Method = methodArr[0];
                }
                var apiUrlArr = item.APIUrl.Split("\"");
                if (apiUrlArr != null && apiUrlArr.Length > 2)
                {
                    item.APIUrl = apiUrlArr[1];
                }
                item.SystemName = systemName;
                item.ServiceFileName = fileName;
            }

            return list;
        }

        static List<ModelAPIDescription> ReadAPIDescrption(string fileName)
        {
            var list = ExcelTo.ImportExcelToList<ModelAPIDescription>(fileName);

            if (list == null || list.Count == 0)
            {
                Console.WriteLine($"{fileName}文件内容为空，或没获取到数据");
                return null;
            }

            list = list.Where(m =>
            !string.IsNullOrEmpty(m.API) &&
            !string.IsNullOrEmpty(m.Description) &&
            m.API != "API" &&
            m.Description != "Description").Distinct().ToList();

            Console.WriteLine(">>>有效接口为：" + (list?.Count ?? 0));
            foreach (var item in list)
            {
                var arrays = item.API.Split(' ').ToArray();
                if (arrays == null || arrays.Length == 0) continue;
                item.Type = arrays[0];
                item.API = arrays[1];
            }

            return list;
        }
    }
}
