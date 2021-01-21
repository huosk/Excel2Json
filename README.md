# Excel 导出工具

支持 Lua 脚本自定导出格式

定义导出脚本`main.lua`

```lua
function export(context)

    local tb_devices =  context:parse_excel_sheet("devices.xls",0,2,3)
    for i = 1,#tb_devices do
        as_int(tb_devices[i],"id")
        as_int(tb_devices[i],"type")
    end

    local tb_device_ids = {}
    for i = 1,#tb_devices do
        table.insert(tb_device_ids, 1, tb_devices[i].id)
        as_int(tb_device_ids,i)
    end

    context:add_exporter(context.Json)

    context:save_doc("devices_ids",tb_device_ids)
    context:save_doc("devices",tb_devices)
end

return export
```

导出

```bat
ExcelExporter -i main.lua
```

详细参数可以查看帮助

```bat
ExcelExporter --help
```