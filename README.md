# Excel2Json

Excel 转 Json 工具，支持基本数据类型：

- string
- int
- float
- boolean

同时支持字符分割数组，声明格式如下：

type[*splitChar*]

type可以是支持的任意数据类型，`splitChar` 是用于分割字符串的符号，可以是任意符号，通常为逗号(，)，竖线（|）等。

## Excel 格式要求

使用此工具对 Excel 有一定的要求，主要以下几点：

1. Excel 表由四部分组成：
   - 标题行：指定列数据的含义
   - 类型行：指定列的数据类型
   - 字段行：生成Json 时的属性名称
2. 需要在 Excel 表的第一行一列，指定单元格标题行的行号