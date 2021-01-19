using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using XLua;
using ExcelExporter.Core;

namespace ExcelExporter
{
    class Program
    {
        static void Main(string[] args)
        {
            string luaFile = args[0];

            LuaEnv luaEnv = new LuaEnv();
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

            luaEnv.Dispose();
        }
    }
}
