# Excel to Word 模板填充工具

这是一个跨平台的开源 Python 工具，用于：

- 读取 Excel (`.xlsx`) 表格数据
- 将指定列映射到 Word (`.docx`) 模板占位符
- 自动生成多个 Word 文档
- 支持按字段分组生成一个 Word 文档
- 支持表格模板行复制与批量填充
- 提供命令行和 Tkinter GUI 两种使用方式

## 新特性

- 支持多种占位符格式：`{{KEY}}`、`<<KEY>>`、`[[KEY]]`
- 支持按 `--group-by` 字段分组输出一个 Word 文档
- 支持表格模板行复制标记，例如 `{{#repeat}}`
- 提供 `excel2word` 和 `excel2word-gui` 命令

## 安装

```bash
pip install -r requirements.txt
```

## 命令行用法

```bash
python -m excel2word \
  --excel data.xlsx \
  --template template.docx \
  --output generated \
  --map 盒号=BOX_NO \
  --map 文号=DOC_NO \
  --map 责任者=OWNER \
  --map 文件题名=TITLE \
  --output-name "{BOX_NO}_{DOC_NO}.docx"
```

或使用入口脚本：

```bash
excel2word --excel data.xlsx --template template.docx --output generated --map 盒号=BOX_NO --map 文号=DOC_NO
```

Word 模板中可以使用如下占位符：

```text
盒号：{{BOX_NO}}
文号：{{DOC_NO}}
责任者：{{OWNER}}
标题：{{TITLE}}
```

## 分组输出

如果你希望按某个字段（例如 `盒号`）生成一个 Word 文档，可以使用：

```bash
excel2word --excel data.xlsx --template template.docx --output generated --group-by 盒号 --output-name "{BOX_NO}.docx" --map 盒号=BOX_NO --map 文号=DOC_NO
```

模板可以在文档正文中引用分组字段：

```text
盒号：{{BOX_NO}}
共计：{{GROUP_ROW_COUNT}} 条记录
```

## 表格行复制

在 Word 模板表格中，插入一行作为数据模板，并在某个单元格中添加 `{{#repeat}}` 标记。工具会复制该行，并替换对应字段。

例如：

| 序号 | 档号 | 文号 | 责任者 | 题名 | {{#repeat}} |
|------|------|------|--------|------|-----------|
| {{ROW_INDEX}} | {{BOX_NO}} | {{DOC_NO}} | {{OWNER}} | {{TITLE}} | {{#repeat}} |

然后在命令行中使用 `--group-by 盒号` 生成一个整表格输出文档。

## JSON 配置示例

```json
{
  "mappings": {
    "盒号": "BOX_NO",
    "文号": "DOC_NO",
    "责任者": "OWNER",
    "文件题名": "TITLE"
  },
  "output_name": "{BOX_NO}.docx",
  "group_by": "盒号",
  "repeat_marker": "#repeat"
}
```

运行：

```bash
excel2word --excel data.xlsx --template template.docx --output generated --config config_example.json
```

## GUI 使用

运行：

```bash
excel2word-gui
```

GUI 允许你选择 Excel、Word 模板、输出目录，并输入列映射、分组字段和输出文件名模板。

## 注意事项

- 模板占位符建议不要跨多个 run 拆分，否则替换效果可能不稳定。
- 输出目录会自动创建。
- 若输出文件名称冲突，工具会自动追加后缀以避免覆盖。

## 开发

```bash
python -m pip install -r requirements.txt
python -m excel2word --help
```
