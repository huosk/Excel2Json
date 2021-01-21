using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using XLua;
using ExcelExporter.Core;
using CommandLine;
using CommandLine.Text;

namespace ExcelExporter
{
    class Options
    {
        [Option('i', "input", Required = true, HelpText = "lua export script list.")]
        public IEnumerable<string> InputFiles { get; set; }

        [Usage]
        public static IEnumerable<Example> Examples
        {
            get
            {
                yield return new Example("export single lua", new Options() { InputFiles = new string[] { "main.lua" } });
                yield return new Example("export multi lua", new Options() { InputFiles = new string[] { "exp.lua", "exp2.lua", "exp3.lua" } });
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args)
                .WithParsed(RunWithOption)
                .WithNotParsed(HandlerErrorArgument);
        }

        static void RunWithOption(Options options)
        {
            LuaEnv luaEnv = new LuaEnv();
            LuaAPI.init_api(luaEnv);

            foreach (var luaFile in options.InputFiles)
            {
                byte[] chunk = File.ReadAllBytes(luaFile);
                object[] objs = luaEnv.DoString(chunk);
                foreach (LuaFunction func in objs)
                {
                    if (func == null)
                        continue;

                    string luaFullPath = Path.GetFullPath(luaFile);
                    var ctx = new ExportContext(luaEnv, Path.GetDirectoryName(luaFullPath));
                    func.Action(ctx);
                }
            }

            luaEnv.Dispose();
        }

        static void HandlerErrorArgument(IEnumerable<Error> errors)
        {
            foreach (var err in errors)
            {
                Console.Error.WriteLine($"{err}");
            }
        }
    }
}
