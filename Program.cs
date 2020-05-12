﻿using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelToJson
{
    class Program
    {
        static void Main(string[] args)
        {
            AppConfig config = LoadConfigure();

            string srcDir = Path.GetFullPath(config.sources);

            if (!Directory.Exists(srcDir))
            {
                Console.Error.WriteLine("指定的Excel文件地址不存在");
                return;
            }

            int successCount = 0;
            int failedCount = 0;
            string[] files = Directory.GetFiles(srcDir, "*.xls?", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
                string file = files[i];
                try
                {
                    ExcelFile excelFile = new ExcelFile(file);

                    var rows = excelFile.GetRows();
                    JObject jObject = new JObject();
                    JArray jArray = new JArray();
                    jObject["0"] = jArray;

                    for (int rIndex = 0; rIndex < rows.Count; rIndex++)
                    {
                        var row = rows[rIndex];
                        JObject cellObj = new JObject();
                        for (int colIndex = 0; colIndex < row.Length; colIndex++)
                        {
                            var cell = row[colIndex];
                            cellObj.Add(new JProperty(cell.Key, cell.Value));
                        }

                        jArray.Add(cellObj);
                    }

                    string json = jObject.ToString();
                    string destDir = Path.GetFullPath(config.destination);
                    string relFilePath = Path.GetRelativePath(srcDir, file);

                    string relDestFile = Path.ChangeExtension(relFilePath, ".json");
                    string destFile = Path.GetFullPath(relDestFile, destDir);

                    string destFolder = Path.GetDirectoryName(destFile);
                    if (!Directory.Exists(destFolder))
                        Directory.CreateDirectory(destFolder);

                    File.WriteAllText(destFile, json);

                    successCount++;
                    Console.WriteLine("生成成功::{0}", destFile);
                }
                catch (Exception e)
                {
                    failedCount++;
                    Console.Error.WriteLine("生成失败，文件：{0}，错误：{1}", file, e.Message);
                }
            }

            Console.WriteLine(">>>>>>>>>>>>>>>>>>>>> 生成完成 <<<<<<<<<<<<<<<<<<<<<<<<");
            Console.WriteLine("                  成功：{0}；失败{1}", successCount, failedCount);
            Console.WriteLine("按任意键退出");
            Console.ReadKey();
        }

        static AppConfig LoadConfigure()
        {
            try
            {
                string json = File.ReadAllText("config.json");
                return JsonConvert.DeserializeObject<AppConfig>(json);
            }
            catch (Exception e)
            {
                Console.Error.WriteLine("配置文件加载错误::{0}", e.Message);
                throw;
            }
        }
    }

    public class AppConfig
    {
        public string sources = "..\\";
        public string destination = "..\\generate";
    }
}
