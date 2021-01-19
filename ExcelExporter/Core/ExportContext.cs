using ExcelExporter.Exportor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLua;

namespace ExcelExporter.Core
{
    public class ExportContext
    {
        public string workDirectory;
        public string src;
        public string output;

        public bool configeReady;

        private LuaEnv luaEnv;
        private List<string> includes;
        private List<IExporter> exporters;

        public IExporter Json { get; }
        public IExporter Lua { get; }

        public ExportContext(LuaEnv luaEnv, string workDir)
        {
            this.luaEnv = luaEnv;
            this.workDirectory = workDir;
            this.includes = new List<string>();
            this.exporters = new List<IExporter>();
            this.Json = new JsonExporter();
            this.Lua = new LuaExporter();
        }

        public void config(LuaTable tb)
        {
            string dest = tb.GetInPath<string>("output");

            this.src = src ?? "";
            this.output = dest ?? "";

            configeReady = true;
        }

        public LuaTable parse_excel_sheet(string file, int sheetNum, int titleNum, int dataNum)
        {
            string fileFullPath = GetFullPath(file);
            var list = ExcelFile.ParseFile(fileFullPath, sheetNum, titleNum - 1, dataNum - 1);
            LuaTable tb = luaEnv.NewTable();
            for (int i = 1; i <= list.Count; i++)
            {
                Dictionary<string, object> dic = list[i - 1];
                if (dic == null)
                    continue;

                LuaTable row = luaEnv.NewTable();
                foreach (var kvp in dic)
                {
                    row.Set(kvp.Key, kvp.Value);
                }

                tb.Set(i, row);
            }

            return tb;
        }

        public List<Dictionary<string, object>> parse_excel_sheet(string file, string sheet, int titleNum, int dataNum)
        {
            string fileFullPath = GetFullPath(file);
            return ExcelFile.ParseFile(fileFullPath, sheet, titleNum - 1, dataNum - 1);
        }

        public void add_exporter(IExporter exp)
        {
            if (!exporters.Contains(exp))
                exporters.Add(exp);
        }

        public void save_doc(string file, LuaTable tb)
        {
            string fileFull = GetFullPath(file);

            for (int i = 0; i < exporters.Count; i++)
            {
                var exp = exporters[i];
                if (exp == null)
                    return;

                exp.export_to_file(fileFull, tb);
            }
        }

        /// <summary>
        /// 获取
        /// </summary>
        /// <returns></returns>
        private string GetFullPath(string file)
        {
            if (!System.IO.Path.IsPathRooted(file))
            {
                var combined = System.IO.Path.Combine(workDirectory, file);
                return System.IO.Path.GetFullPath(combined);
            }
            else
            {
                return System.IO.Path.GetFullPath(file);
            }
        }
    }
}
