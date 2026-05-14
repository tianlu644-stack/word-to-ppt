# Word to PPT 转换工具

将 Word 文档（.docx）自动转换为 PowerPoint 演示文稿（.pptx）。

## 功能特性

- ✅ 自动读取 Word 文档内容
- ✅ 智能识别标题和内容
- ✅ 保留文本格式和结构
- ✅ 生成专业美观的 PPT 幻灯片
- ✅ 支持批量转换
- ✅ 自定义样式支持

## 环境要求

- Python 3.7+
- 依赖包详见 `requirements.txt`

## 安装

### 1. 克隆或进入仓库

```bash
git clone https://github.com/tianlu644-stack/word-to-ppt.git
cd word-to-ppt
```

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 方法 1: 转换单个文件

```bash
python word_to_ppt.py 你的文件.docx
```

或指定输出文件名：

```bash
python word_to_ppt.py 你的文件.docx 输出文件.pptx
```

### 方法 2: 批量转换所有 Word 文件

```bash
python convert.py
```

这会自动转换当前目录下的所有 `.docx` 文件，输出到 `output/` 目录。

## 文件说明

| 文件 | 说明 |
|------|------|
| `word_to_ppt.py` | 主转换模块，包含 `WordToPptConverter` 类 |
| `convert.py` | 批量转换脚本 |
| `requirements.txt` | Python 依赖配置 |
| `.gitignore` | Git 忽略文件配置 |

## 转换流程

1. **读取 Word 文档** - 提取所有段落文本
2. **智能分析** - 识别标题和内容段落
3. **生成幻灯片** - 创建对应的 PPT 幻灯片
4. **样式应用** - 应用专业的配色和排版
5. **保存文件** - 输出为 `.pptx` 文件

## 示例

### 输入

一个 Word 文档包含：

```
轨道 6 号线推广方案

项目概览
项目介绍
建设意义

技术方案
技术路线
关键技术

实施计划
时间安排
成本预算
```

### 输出

生成的 PPT 包含：

1. 标题页："轨道 6 号线推广方案"
2. 项目概览页：包含项目介绍和建设意义
3. 技术方案页：包含技术路线和关键技术
4. 实施计划页：包含时间安排和成本预算

## 自定义样式

你可以编辑 `WordToPptConverter` 类中的以下方法来自定义样式：

- `add_title_slide()` - 标题幻灯片样式
- `add_content_slide()` - 内容幻灯片样式

### 修改背景色

```python
fill.fore_color.rgb = RGBColor(255, 0, 0)  # 改为红色
```

### 修改字体大小和颜色

```python
paragraph.font.size = Pt(50)  # 字体大小
paragraph.font.color.rgb = RGBColor(0, 0, 255)  # 蓝色
```

## 常见问题

### Q: 转换后 PPT 格式不对？
A: Word 文档的结构会影响转换结果。确保标题和内容的分层清晰。

### Q: 能否保留 Word 中的图片？
A: 当前版本只转换文本。图片支持会在后续版本添加。

### Q: 如何修改输出 PPT 的样式？
A: 编辑 `word_to_ppt.py` 中的 `add_title_slide()` 和 `add_content_slide()` 方法。

## 更新日志

### v1.0.0 (2026-05-14)
- ✨ 首个版本发布
- ✅ 基本 Word 到 PPT 转换功能
- ✅ 批量转换支持
- ✅ 自定义样式支持

## 贡献

欢迎提交 Issue 和 Pull Request！

## 许可证

MIT License
