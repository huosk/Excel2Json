using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelExporter.Core;
using XLua;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelExporter.Exportor
{
    public class JsonExporter : IExporter
    {
        public void export_to_file(string fileNameWithoutExt, LuaTable table)
        {
            string file = Path.ChangeExtension(fileNameWithoutExt, "json");


            JContainer jObject = null;
            parse_lua(table, ref jObject);

            File.WriteAllText(file, jObject.ToString(), Encoding.UTF8);
        }

        private bool is_array(LuaTable table)
        {
            var keys = table.GetKeys();
            foreach (var key in keys)
            {
                if (!typeof(long).IsAssignableFrom(key.GetType()))
                {
                    return false;
                }
            }

            return true;
        }

        private void parse_lua(LuaTable table, ref JContainer container)
        {
            if (table == null)
                return;

            if (is_array(table))
            {
                if (container == null)
                    container = new JArray();

                parse_lua_array(table, (JArray)container);
            }
            else
            {
                if (container == null)
                    container = new JObject();

                parse_lua_object(table, (JObject)container);
            }
        }

        private void parse_lua_array(LuaTable table, JArray array)
        {
            var keys = table.GetKeys<long>().ToArray();
            foreach (var key in keys)
            {
                var val = table[key];
                if (val == null)
                {
                    continue;
                }

                if (val.GetType() == typeof(LuaTable))
                {
                    JContainer c = null;
                    parse_lua((LuaTable)val, ref c);
                    array.Add(c);
                }
                else
                {
                    array.Add(JToken.FromObject(val));
                }
            }
        }

        private void parse_lua_object(LuaTable table, JObject obj)
        {
            var keys = new List<object>();
            foreach (var k in table.GetKeys())
            {
                keys.Add(k);
            }

            foreach (var item in keys)
            {
                string jsonKey = item.ToString();
                var val = table[item];
                if (val == null)
                {
                    obj.Add(jsonKey, null);
                }
                else
                {
                    if (val.GetType() == typeof(LuaTable))
                    {
                        JContainer nobj = null;
                        parse_lua((LuaTable)val, ref nobj);
                        obj.Add(jsonKey, nobj);
                    }
                    else
                    {
                        obj.Add(jsonKey, JToken.FromObject(val));
                    }
                }
            }
        }
    }
}
