# 通用公文排版 Skill

把文字整理成排版规范的 Word 文档，也可以把已有的 DOCX 重新排版。它只负责结构识别和排版，不会擅自改写内容。你只需要把文字或文件交给 Codex，并说明想做什么，不需要会写代码。

## 先看答案

| 你可能最关心的问题 | 答案 |
|---|---|
| 能做什么？ | 根据纯文本或 Markdown 生成 DOCX；检查并重新排版已有 DOCX。 |
| 目前支持什么平台？ | V1 面向 Windows Codex。其他系统尚未做完整适配和验收。 |
| 会改坏原文件吗？ | 不会。已有文档始终另存为新文件，原件保持不变。 |
| 需要会代码吗？ | 不需要。安装后直接和 Codex 对话即可。 |
| 必须安装 Microsoft Word 吗？ | 不需要。核心排版由 Python 直接完成；安装 Word 或 LibreOffice 后，才可以增加视觉预览检查。 |
| 首次使用还要装东西吗？ | Skill 会先检查 Python 和所需组件。发现缺少时会告诉你并等待确认，不会自行安装。 |
| 第一次使用要填很多参数吗？ | 不用。第一次直接采用内置通用格式。 |
| 能记住我单位的格式吗？ | 可以。只有你主动修改并发送“确认保存”后，个性化格式才会长期生效。 |
| 能处理哪些文件？ | 新建或处理 `.docx`。`.doc`、`.docm` 需要先另存为 `.docx`。 |
| 表格、图片和文本框会怎样？ | 表格统一字体族、黑色文字和全局加粗设置，其余版式保留；图片和文本框原样保留并提醒人工检查。 |
| 字体没装齐怎么办？ | 自动尝试替代字体并给出提示，不会偷偷下载字体。 |
| 文档内容会写进长期配置吗？ | 不会。长期配置只保存排版参数，不保存正文、姓名、单位、日期或文件路径。 |
| 会把文档上传到其他地方吗？ | 这个 Skill 的处理脚本在本机运行，不含上传功能；Codex 服务本身如何处理数据，请以你所用版本和账号的相关设置为准。 |

最常见的使用过程很简单：提供文字或 DOCX → 确认识别结果 → 获得一个新 DOCX。

它适合日常文字材料的统一排版，不适合用来写作正文、识别扫描件、转换 PDF，或重新设计复杂表格和图文版面。

## 30 秒安装

### 方式一：一行命令安装

在 PowerShell 或终端中运行：

```powershell
npx -y skills add Triwood-79/official-document-formatting -g -a codex -y --copy
```

安装完成后，新建一个 Codex 任务即可使用。

这条命令需要电脑已经安装 Node.js。仓库为 Private 时，安装者还需要拥有仓库访问权限并在本机登录 GitHub；仓库改为 Public 后可直接安装。

### 方式二：把这段话直接发给 Codex

不想自己操作命令行，可以复制下面整段文字发给有本机文件和终端权限的 Codex：

```text
请从 https://github.com/Triwood-79/official-document-formatting 安装“通用公文排版”Skill。
先只读检查仓库根目录的 README.md 和 SKILL.md，确认 Skill 名称是 official-document-formatting。
如果本机已经存在同名 Skill，请停止安装并告诉我现状，不要覆盖。
确认无冲突后，运行：npx -y skills add Triwood-79/official-document-formatting -g -a codex -y --copy
不要自动安装 Python 依赖或字体。安装完成后，请检查 Skill 是否已经进入 Codex 全局技能目录，
并告诉我实际安装路径和检查结果。
```

### 方式三：先检查，再安装

如果是单位电脑，或你希望先看清仓库内容再决定是否安装，建议采用这种方式：

1. 打开 [GitHub 仓库](https://github.com/Triwood-79/official-document-formatting)，先阅读 `README.md` 和 `SKILL.md`。
2. 下载 ZIP，解压到一个临时文件夹，不要直接放进 Codex 技能目录。
3. 请 Codex 检查文件清单，运行隐私扫描和自动化测试。
4. 检查通过后，再把整个文件夹复制到 Codex 可识别的个人 Skill 目录。当前官方文档列出的用户级目录是 `%USERPROFILE%\.agents\skills\official-document-formatting`；使用 Codex 内置 Skill Installer 时，也可能安装到 `%CODEX_HOME%\skills\official-document-formatting`（未设置 `CODEX_HOME` 时通常位于 `%USERPROFILE%\.codex\skills`）。以安装工具返回的实际路径为准，不要同时维护两份可变源代码。
5. 新建一个 Codex 任务，用“查看当前排版格式”验证是否安装成功。

也可以把下面这段话发给 Codex，让它在每一步完成后向你报告：

```text
请用“先检查、后安装”的方式安装这个 Skill：
https://github.com/Triwood-79/official-document-formatting

先把仓库下载到临时检查目录，不要直接安装，也不要覆盖任何同名目录。
依次检查 README.md、SKILL.md 和文件清单，运行 scripts/privacy_scan.py，
再运行 python -m unittest discover -s tests -v。
把检查结果告诉我；只有检查通过后，才把 official-document-formatting 文件夹复制到
Codex 全局技能目录。不要自动安装依赖或字体，缺少环境时先说明缺什么并等待我确认。
安装完成后，请告诉我实际路径，并验证新任务能识别这个 Skill。
```

这种方式多一步，但能在安装前确认文件、测试结果和目标位置，也避免把网上获取的命令直接交给 PowerShell 执行。

## 安装后怎么用

### 把文字生成 Word 文档

把内容粘贴给 Codex，然后说：

```text
使用 $official-document-formatting，把下面的内容生成排版后的 Word 文档。
```

示例内容：

```text
示例标题
某单位
某年某月某日

一、一级标题
这里是正文内容。

（一）二级标题
这里是下一段正文内容。
```

没有提供的署名、单位、日期或事实内容，Skill 不会自行补写。V1 不新建或重构完整的版记结构；已有印发机关和印发日期行只统一为正文字体族和黑色，原字号、对齐、分隔线及位置保持不变，并提醒你手动完成复杂布局。

### 重新排版已有 Word 文档

把 `.docx` 文件交给 Codex，然后说：

```text
使用通用公文排版 Skill，检查并重新排版这份文档。
```

Skill 会先告诉你识别到了哪些主标题、标题、正文和说明内容。你确认或纠正后，它才会开始排版，并生成：

```text
原文件名_排版后.docx
```

如果重名，会自动增加序号。原文件不会被覆盖。

### 临时改一次格式

直接在当前请求里说明即可，例如：

```text
这一次改成单面打印，页码放在页面下方居中，不要保存为长期设置。
```

这类要求只影响当前文档。

## 想改成自己的固定格式

第一次使用直接采用内置格式，不会弹出问卷。需要修改时，对 Codex 发送下面的固定口令。

| 你想做什么 | 发送的口令 |
|---|---|
| 查看现在使用的格式 | `查看当前排版格式` |
| 修改部分参数 | `修改排版格式` |
| 重新填写全部参数 | `重新录入排版格式` |
| 保存修改草稿 | `确认保存` |
| 放弃本次修改 | `取消修改` |
| 恢复内置格式 | `恢复通用公文预设格式` |

修改时，Codex 会按页边距与打印、字体、标题、正文、说明和页码等内容分组提问，并在每个问题旁显示当前值。固定行距是 V1 的明确提问项，会分别确认主标题、一级标题、二级标题、正文、补充参考资料和说明内容的行距；也可以直接回答“除主标题外统一为 30 pt”这类合并设置。

发送 `重新录入排版格式` 后，Codex 会先按组展示整套当前值，再逐项询问全部 V1 开放字段。回答“保持当前值”的项目也会明确写入新草稿，不会因为重新录入而悄悄恢复成内置默认值。V1 不开放版记、纸张尺寸、页面方向和页码装饰位置的修改。保存前会展示“当前值 → 新值”的差异；没有收到 `确认保存`，原来的格式就继续生效。

排版参数按这个顺序决定：当前文档的临时要求 > 已确认保存的个人格式 > 内置通用格式。

<details>
<summary>查看内置通用格式</summary>

| 对象 | 默认设置 |
|---|---|
| 页面 | A4 纵向 |
| 页边距 | 上 3.7 cm、下 3.5 cm、左 2.8 cm、右 2.6 cm |
| 全局字形 | 全部加粗，包括页码 |
| 主标题 | 方正小标宋_GBK，二号 22 pt，居中，固定行距 35 pt |
| 一级标题 | 方正黑体_GBK，三号 16 pt，固定行距 30 pt |
| 二级标题 | 方正楷体_GBK，三号 16 pt，固定行距 30 pt |
| 正文 | 方正仿宋_GBK，三号 16 pt，首行缩进 2 字符，固定行距 30 pt，两端对齐 |
| 数字和英文 | Times New Roman，三号 16 pt |
| 补充参考资料 | 方正楷体_GBK，小三号 15 pt |
| 说明内容 | 方正楷体_GBK，三号 16 pt，居中，固定行距 35 pt |
| 打印模式 | 默认双面打印 |
| 页码 | 页面下方外侧，奇数页右、偶数页左；宋体、四号 14 pt、粗体、半角阿拉伯数字，格式为 `— 1 —` |

选择单面打印或平板阅读模式后，页码会改为页面下方居中。

</details>

## 使用前要知道的边界

- 只直接处理 `.docx`。旧版 `.doc` 和带宏的 `.docm` 请先另存为 `.docx`。
- 带批注、修订、加密或文档保护的文件，需要先提供一份已处理完成的干净副本。
- 排版后的普通文字统一为黑色，不再保留原文中的蓝色、红色或主题色。
- 表格会统一字体族、文字颜色和全局加粗设置：表题用方正小标宋，第一层表头用方正黑体，第二层表头用方正楷体，表格正文用方正仿宋；字号、对齐、行距、边框、底纹、行高和列宽保持原样。执行前会展示表题候选和表头层数供确认。
- 图片、文本框和特殊横向页面会保留，但不自动统一内部格式。
- V1 不新建或重构完整的版记结构；已有印发机关和印发日期行仅统一为方正仿宋、Times New Roman 和黑色，原字号、粗体、对齐、段距、分隔线及位置保持不变，复杂布局仍需人工完成。
- Skill 负责识别结构和排版，不负责补写事实或推测缺失信息。
- 输出后会清理常见的作者、最后修改者、公司和自定义属性等文档元数据。
- 没有可用的 Word 或 LibreOffice 时，只能确认文档结构通过检查，不能把它表述为已经完成视觉验收。

## 字体不齐时会发生什么

Skill 会先使用首选字体。电脑没有对应字体时，按下面的顺序尝试替代：

- 方正小标宋_GBK → 方正小标宋简体
- 方正黑体_GBK → 黑体
- 方正楷体_GBK → 楷体_GB2312
- 方正仿宋_GBK → 仿宋_GB2312
- 宋体 → SimSun

如果首选和替代字体都没有，文档仍然可以生成，但 Codex 会明确提醒你：当前电脑无法完成字体效果的视觉检查。Skill 不会自动下载或捆绑字体。

## 常见问题

### 我已经安装过了，怎样更新？

运行：

```powershell
npx -y skills update official-document-formatting -g -y
```

更新前如果改过本地 Skill 文件，建议先备份或提交到自己的 Git 仓库，避免产生冲突。

### 换一台电脑还能使用原来的设置吗？

重新安装 Skill 后可以继续使用它，但个人排版配置默认保存在当前电脑，不会跟着 Git 仓库自动同步。需要在新电脑重新录入，或由你明确决定复制配置文件。

### 为什么我看到了“结构验证通过”，却没有看到“视觉验收通过”？

结构验证检查的是页边距、字体名称、字号、段落和页码字段等文档内部设置。视觉验收还需要 Word 或 LibreOffice 把文档真正渲染出来。缺少渲染程序时，Skill 不会把结构检查说成肉眼效果已经完全正确。

### 为什么不直接覆盖原文档？

自动识别可能需要人工纠正，复杂 Word 元素也可能有特殊设置。另存新文件方便对照，也能随时回到原件。

<details>
<summary>给开发者和高级用户：环境、脚本与测试</summary>

核心环境：Python 3、`python-docx`、`lxml`。不要求安装 Microsoft Word。

检查环境：

```powershell
.\scripts\preflight.ps1
```

经用户确认后安装缺少的 Python 依赖：

```powershell
python -m pip install -r requirements.txt
```

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

分类文件默认保留。用户明确要求删除时，先做只读校验：

```powershell
python scripts\cleanup_classification.py `
  --classification classification.json `
  --input input.docx `
  --output input_排版后.docx
```

展示校验结果并再次确认后，使用上一步返回的 SHA-256 删除；文件一旦变化就会拒绝：

```powershell
python scripts\cleanup_classification.py `
  --classification classification.json `
  --input input.docx `
  --output input_排版后.docx `
  --confirmed `
  --expected-sha256 <上一步返回的 SHA-256>
```

运行自动化测试：

```powershell
python -m unittest discover -s tests -v
```

运行发布前隐私扫描：

```powershell
python scripts\privacy_scan.py .
```

</details>

## 许可证

本项目使用 [MIT License](LICENSE)。

## 反馈

如果遇到结构识别、排版结果或安装问题，可以在 [GitHub Issues](https://github.com/Triwood-79/official-document-formatting/issues) 中提交。请先移除文档中的真实业务内容，只提供可以公开的复现示例。
