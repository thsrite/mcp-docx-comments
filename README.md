# mcp-docx-comments

将 Word 文档（.docx）中的批注（comments）内联到正文中的 MCP Server。

批注会以 **【批注：内容】** 的格式插入到正文对应位置，红色加粗显示。

## 工具列表

### inline_comments_base64

通过 base64 编码传输文件，适用于远程调用场景。

**参数：**
- `file_base64`（必填）：base64 编码的 .docx 文件内容
- `filename`（可选）：文件名，默认 `document.docx`

**返回：** JSON 字符串，包含：
- `filename`：输出文件名
- `file_base64`：处理后的 base64 编码文件
- `total_comments`：批注总数
- `processed_comments`：已处理批注数

### inline_comments_file

处理本地文件路径，适用于本地调用场景。

**参数：**
- `input_path`（必填）：输入的 .docx 文件路径
- `output_path`（可选）：输出文件路径，默认在原文件名后加 `_批注内联`

**返回：** 处理结果信息

## 安装与使用

### 通过 mcp-gateway 部署

在 mcp-gateway Web UI 中添加服务器，GitHub URL 填入：

```
https://github.com/thsrite/mcp-docx-comments.git
```

网关会自动安装依赖并启动服务，通过 HTTP 调用即可使用。

### 本地运行

```bash
pip install -e .
mcp-docx-comments
```
