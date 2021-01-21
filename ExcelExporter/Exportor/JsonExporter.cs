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
                if (is_internal_key(key as string))
                    continue;

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
            Dictionary<object, Type> __type_map__ = null;
            table.Get<string, Dictionary<object, Type>>("__type_map__", out __type_map__);
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
                    if (__type_map__ != null && __type_map__.TryGetValue(key, out Type tp))
                    {
                        array.Add(JToken.FromObject(convert(val, tp) ?? val));
                    }
                    else
                    {
                        array.Add(JToken.FromObject(val));
                    }
                }
            }
        }

        private object convert(object val, Type tp)
        {
            if (tp == typeof(int) || tp == typeof(long))
            {
                return JToken.FromObject(Convert.ToInt64(val));
            }
            else if (tp == typeof(string))
            {
                return JToken.FromObject(Convert.ToString(val));
            }
            else if (tp == typeof(float) || tp == typeof(double))
            {
                return JToken.FromObject(Convert.ToDouble(val));
            }
            else if (tp == typeof(bool))
            {
                return JToken.FromObject(Convert.ToBoolean(val));
            }
            else
            {
                return null;
            }
        }

        private bool is_internal_key(string key)
        {
            return !string.IsNullOrEmpty(key) && key.StartsWith("__");
        }

        private void parse_lua_object(LuaTable table, JObject obj)
        {
            var keys = new List<object>();
            foreach (var k in table.GetKeys())
            {
                keys.Add(k);
            }

            Dictionary<object, Type> __type_map__ = null;
            table.Get<string, Dictionary<object, Type>>("__type_map__", out __type_map__);
            foreach (var item in keys)
            {
                string jsonKey = item.ToString();
                if (is_internal_key(jsonKey))
                    continue;

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
                        if (__type_map__ != null && __type_map__.TryGetValue(jsonKey, out Type tp))
                        {
                            obj.Add(jsonKey, JToken.FromObject(convert(val, tp) ?? val));
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
}
