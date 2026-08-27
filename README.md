# 通用公文排版 Skill

一个用于生成和重新排版中文公文 DOCX 文件的 Codex Skill。它内置了一套通用排版预设，也允许用户通过确认式问答保存自己的排版参数。

## 主要功能

- 根据纯文本或 Markdown 生成排版后的 `.docx` 文档。
- 自动识别已有 DOCX 中的主标题、一级标题、二级标题、正文、说明内容和补充参考资料。
- 排版已有文档前先展示识别摘要，用户确认后才执行。
- 始终另存新文件，不覆盖原文档。
- 支持双面打印页码和单面/平板阅读页码。
- 支持字体检测、字体回退和缺失字体提示。
- 支持查看、修改、取消、恢复和回退个人排版配置。
- 清理输出 DOCX 中的作者、最后修改者、公司和自定义属性等元数据。
- 保留表格、图片和文本框等复杂元素，不擅自改动其内部内容。

## 适用范围

- Windows Codex。
- 新建 `.docx` 文档。
- 重新排版已有 `.docx` 文档。

暂不直接处理 `.doc` 和 `.docm`。请先将这类文件另存为 `.docx`。带有批注、修订、加密或文档保护的文件，需要先生成一份干净副本。

## 安装

将整个 `official-document-formatting` 文件夹放入 Codex 技能目录：

```text
%CODEX_HOME%\skills\official-document-formatting
```

如果没有设置 `CODEX_HOME`，可使用：

```text
%USERPROFILE%\.codex\skills\official-document-formatting
```

重新打开任务后，即可通过 `$official-document-formatting` 显式调用，也可以直接提出生成或排版公文 DOCX 的需求。

## 环境检查

Skill 使用 Python 直接处理 DOCX，不要求安装 Microsoft Word。首次使用前可以运行：

```powershell
.\scripts\preflight.ps1
```

核心依赖为：

- Python 3
- `python-docx`
- `lxml`

如果缺少依赖，请在确认后安装：

```powershell
python -m pip install -r requirements.txt
```

Skill 不会静默安装软件，也不会自动下载字体。

## 快速使用

### 根据内容生成 Word 文档

可以直接对 Codex 说：

```text
使用通用公文排版 Skill，把下面的内容生成 DOCX：

示例标题
某单位
某年某月某日
一、一级标题
这里是正文内容。
```

也可以明确调用：

```text
使用 $official-document-formatting 将这段内容生成排版后的 Word 文档。
```

Skill 不会自动编造单位、署名、日期、版记或事实内容。缺少的内容需要由用户提供。

### 重新排版已有 DOCX

上传或指定一个 `.docx` 文件，然后说：

```text
使用通用公文排版 Skill 检查并重新排版这份文档。
```

处理流程如下：

1. 检查文件类型、批注、修订和文档保护状态。
2. 识别各段落的结构并展示摘要。
3. 等待用户确认或纠正识别结果。
4. 生成 `<原文件名>_排版后.docx`。
5. 清理元数据并执行结构验证。

表格、图片、文本框及特殊横向分节会原样保留并提示。

## 默认排版格式

| 对象 | 默认设置 |
|---|---|
| 页面 | A4 纵向 |
| 页边距 | 上 3.7 cm、下 3.5 cm、左 2.8 cm、右 2.6 cm |
| 全局 | 全部加粗，包括页码 |
| 主标题 | 方正小标宋_GBK，二号 22 pt，居中，固定行距 35 pt |
| 一级标题 | 方正黑体_GBK，三号 16 pt，固定行距 30 pt |
| 二级标题 | 方正楷体_GBK，三号 16 pt，固定行距 30 pt |
| 正文 | 方正仿宋_GBK，三号 16 pt，首行缩进 2 字符，固定行距 30 pt，两端对齐 |
| 数字和英文 | Times New Roman，三号 16 pt |
| 补充参考资料 | 方正楷体_GBK，小三号 15 pt |
| 说明内容 | 方正楷体_GBK，三号 16 pt，居中，固定行距 35 pt |
| 版记 | 方正仿宋_GBK，三号 16 pt，仅处理用户提供的内容 |
| 默认打印模式 | 双面打印 |
| 页码 | 页面下方外侧，奇数页右、偶数页左；宋体、四号 14 pt、粗体、半角阿拉伯数字，格式为 `— 1 —` |

选择单面打印或平板阅读模式后，页码会改为页面下方居中。

## 修改个人排版格式

第一次使用不会弹出设置问卷，而是直接采用内置预设。只有用户主动发出口令时，才会修改长期配置。

### 查看当前配置

```text
查看当前排版格式
```

### 修改部分参数

```text
修改排版格式
```

Skill 会按页面、字体、标题、正文、说明、版记和页码等分组进行问答。回答完成后，它会展示“当前值 → 新值”的差异。

只有用户发送下面的口令后，草稿才会正式生效：

```text
确认保存
```

### 重新录入全部参数

```text
重新录入排版格式
```

### 放弃本次修改

```text
取消修改
```

### 恢复内置预设

```text
恢复通用公文预设格式
```

单次任务中提出的临时格式要求只影响当前文档，不会更新长期配置。

## 字体回退

当首选字体不存在时，Skill 按以下规则尝试替代字体：

- 方正小标宋_GBK → 方正小标宋简体
- 方正黑体_GBK → 黑体
- 方正楷体_GBK → 楷体_GB2312
- 方正仿宋_GBK → 仿宋_GB2312
- 宋体 → SimSun

如果首选字体和替代字体都不存在，Skill 仍会把首选字体名称写入 DOCX，并提示当前设备无法完成该字体的视觉验收。

## 配置优先级

排版参数按以下顺序生效：

1. 当前任务明确提出的临时要求。
2. 用户确认保存的个性化格式。
3. 内置通用公文预设。

用户配置只保存格式参数，不保存姓名、单位、日期、正文或文档路径。

## 输出与隐私

- 已有文档始终另存，不覆盖原件。
- 新建文档从空白 DOCX 生成，不使用带来源信息的真实模板。
- 输出后自动清理可识别作者或设备来源的文档元数据。
- 没有 Word 或 LibreOffice 时，只报告结构验证结果，不声称已经完成视觉验收。

## 命令行使用

普通用户可以直接通过 Codex 对话使用 Skill。需要调试或自动化时，可以调用脚本。

生成新文档：

```powershell
python scripts\docx_engine.py create --input content.md --output output.docx
```

检查已有文档并生成分类文件：

```powershell
python scripts\inspect_docx.py input.docx --write-map classification.json
```

确认分类后重新排版：

```powershell
python scripts\docx_engine.py format-existing `
  --input input.docx `
  --classification classification.json `
  --confirmed
```

运行自动化测试：

```powershell
python -m unittest discover -s tests -v
```
