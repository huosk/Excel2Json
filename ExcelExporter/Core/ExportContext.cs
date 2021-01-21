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

        private LuaEnv luaEnv;
        private List<IExporter> exporters;

        public IExporter Json { get; }
        public IExporter Lua { get; }

        public ExportContext(LuaEnv luaEnv, string workDir)
        {
            this.luaEnv = luaEnv;
            this.workDirectory = workDir;
            this.exporters = new List<IExporter>();
            this.Json = new JsonExporter();
            this.Lua = new LuaExporter();
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

    public class LuaAPI
    {
        public static void init_api(LuaEnv env)
        {
            env.Global.Set<string, Action<LuaTable, object>>("as_int", as_int);
            env.Global.Set<string, Action<LuaTable, object>>("as_long", as_long);
            env.Global.Set<string, Action<LuaTable, object>>("as_double", as_double);
            env.Global.Set<string, Action<LuaTable, object>>("as_float", as_float);
            env.Global.Set<string, Action<LuaTable, object>>("as_bool", as_bool);
            env.Global.Set<string, Action<LuaTable, object>>("as_string", as_string);
        }

        private static Dictionary<object, Type> create_type_map(LuaTable tb)
        {
            Dictionary<object, Type> __type_map__ = null;
            tb.Get<string, Dictionary<object, Type>>("__type_map__", out __type_map__);
            if (__type_map__ == null)
            {
                __type_map__ = new Dictionary<object, Type>();
                tb.Set<string, Dictionary<object, Type>>("__type_map__", __type_map__);
            }
            return __type_map__;
        }

        public static void as_int(LuaTable tb, object key)
        {
            create_type_map(tb)[key] = typeof(int);
        }

        public static void as_long(LuaTable tb, object key)
        {
            create_type_map(tb)[key] = typeof(long);
        }

        public static void as_double(LuaTable tb, object key)
        {
            create_type_map(tb)[key] = typeof(double);
        }

        public static void as_float(LuaTable tb, object key)
        {
            create_type_map(tb)[key] = typeof(float);
        }

        public static void as_bool(LuaTable tb, object key)
        {
            create_type_map(tb)[key] = typeof(bool);
        }

        public static void as_string(LuaTable tb, object key)
        {
            create_type_map(tb)[key] = typeof(string);
        }
    }
}
