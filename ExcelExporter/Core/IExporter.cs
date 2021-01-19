using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLua;

namespace ExcelExporter.Core
{
    public enum ExportorType
    {
        Json,
        Lua,
    }

    public interface IExporter
    {
        void export_to_file(string fileNameWithoutExt, LuaTable table);
    }
}
