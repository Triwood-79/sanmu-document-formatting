# 格式字段与执行规则

## 内置预设

- A4 纵向；页边距为上 3.7 cm、下 3.5 cm、左 2.8 cm、右 2.6 cm。
- 全文加粗，包括页码。
- 主标题：方正小标宋_GBK，二号 22 pt，居中、无缩进、固定行距 35 pt，段前 35 pt。
- 一级标题：方正黑体_GBK，三号 16 pt，首行缩进 2 字符、固定行距 30 pt、两端对齐。
- 二级标题：方正楷体_GBK，三号 16 pt，首行缩进 2 字符、固定行距 30 pt、两端对齐。
- 正文：方正仿宋_GBK，三号 16 pt，首行缩进 2 字符、固定行距 30 pt、两端对齐；数字和英文使用 Times New Roman。
- 补充参考资料：仅限整段括号备注或用户明确标注的段落；方正楷体_GBK，小三号 15 pt。
- 说明内容：主标题后的署名/单位行与日期行；方正楷体_GBK，三号 16 pt，居中、无缩进、固定行距 35 pt。
- 版记：仅处理用户提供的内容；方正仿宋_GBK，三号 16 pt，置于正文最后一页。
- 默认双面打印：页码在下方外侧，奇数页右、偶数页左。单面或平板模式改为下方居中。
- 页码：宋体，四号 14 pt，粗体，半角阿拉伯数字，显示为 `— 1 —`。

## 自动识别

- 第一段非空文字为主标题。
- `一、`、`二、` 等开头为一级标题。
- `（一）`、`（二）` 等开头为二级标题。
- 主标题后连续的短署名/单位行和日期行可识别为说明内容。
- 整段由中英文括号包围的内容可识别为补充参考资料。
- 其他顶层段落为正文。

已有文档必须先显示按段落索引生成的识别结果。用户可通过修改分类 JSON 将任一段落改为 `main_title`、`heading1`、`heading2`、`body`、`reference_note`、`description`、`colophon` 或 `skip`。

## 修改字段

使用点路径修改草稿，例如：

```powershell
python scripts/profile_manager.py begin
python scripts/profile_manager.py set page.margins_cm.top 4.0
python scripts/profile_manager.py set page.print_mode '"single"'
python scripts/profile_manager.py diff
python scripts/profile_manager.py confirm
```

值按 JSON 解析；字符串需要作为 JSON 字符串传入。只允许修改内置预设中已经存在的格式字段。
