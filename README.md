# PDF 去水印转 Word 工具 (pdf-watermark-remover)

批量去除 PDF 中的**页面级文字水印**，并转换为**版式与图片对应的 Word（.docx）**。
提供图形界面与命令行两种使用方式，纯 Python 实现，Windows / macOS / Linux 均可运行。

## 功能特性

- **自动检测水印**：统计跨页重复出现的文本，自动识别疑似水印
- **无损去水印**：Acrobat 标准水印直接从内容流移除绘制指令，正文文字、图片、矢量元素零影响；普通文字水印采用 redaction 擦除（只擦文字、不碰图片）
- **批量转 Word**：基于 pdf2docx，尽量保留版式、文字与截图
- **结果校验**：转换后自动核对水印是否残留，并在日志中显示每个文件的页数/耗时
- **GUI + CLI** 双模式

## 环境要求

- Python 3.9 及以上
- Windows：安装 Python 时务必勾选 **Add Python to PATH**

## 安装

```bash
git clone https://github.com/Maof2023/pdf-watermark-remover.git
cd pdf-watermark-remover
pip install -r requirements.txt

# 国内网络较慢时可用清华镜像：
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

`requirements.txt` 只声明了 `pdf2docx`，它会自动带上 PyMuPDF、python-docx 等全部依赖。

## 使用方法（图形界面）

```bash
python app.py
```

1. 点击 **选择 PDF 文件…**（可按住 Ctrl 多选）或 **选择文件夹…**（处理整个文件夹的 PDF）
2. 输出目录默认为 `源目录\word`，可通过 **浏览…** 修改
3. 点击 **检测水印** 查看自动识别的候选水印（显示出现页数与占比）；
   一般保持勾选 **开始时自动检测** 即可全自动处理
4. 点击 **开始转换**，进度条与日志实时显示每个文件的转换结果
5. 完成后点击 **打开输出目录** 查看生成的 docx

## 使用方法（命令行）

```bash
# 整个目录，自动检测水印并转换
python core.py "D:\资料\pdfs"

# 指定水印文本与输出目录
python core.py "D:\资料\pdfs" -w "你的水印文本" -o "D:\输出目录"

# 仅检测候选水印（先看检测结果再决定）
python core.py --detect "某文件.pdf"

# 调整检测阈值（默认同一文本出现在 >=50% 页面即视为候选）
python core.py --detect "某文件.pdf" --min-ratio 0.8
```

## 打包为 exe（可选，目标电脑无需安装 Python）

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed --name Pdf2WordCleaner ^
  --collect-all pdf2docx --collect-all pymupdf --collect-all pymupdf_layout app.py
```

产物为 `dist\Pdf2WordCleaner.exe`，双击即可运行（首次启动解压稍慢属正常现象）。
macOS / Linux 用户请把续行符 `^` 换成 `\`。

## 工作原理

1. **检测**：用 PyMuPDF 提取每页文本行，统计"同一文本出现在 >= 阈值比例页面"的短文本作为水印候选
2. **去水印**（两种策略按优先级自动选择）：
   - **Acrobat 标准水印**：水印存放在带 `Private/Watermark` 标记的 Form XObject 中，页面内容流仅删除指向它的 `Do` 绘制指令 —— 完全无损
   - **普通文字水印**：`search_for` 定位后 redaction 擦除（`PDF_REDACT_IMAGE_NONE` 保证只擦文字不碰图片）
3. **转换**：pdf2docx 解析段落/表格/图片并重建 docx

## 已知限制

- 烧在**图片像素内部**的水印属于图片本身，无法无损去除（会原样保留）
- 普通文字水印的 redaction 方案对**横排**文字效果最好，斜排文字可能定位不到
- 复杂表格/图文混排页面的版式可能有轻微偏差（pdf2docx 还原精度所致）
- 自动检测取"出现页数最多"的候选；页眉页脚等固定元素也可能入选，
  建议先用【检测水印】人工确认后再批量处理

## 验证数据

在 18 个中文 PDF（共 370 页，Acrobat 45° 斜排半透明文字水印）上实测：
18/18 转换成功，文字保留率 100%（逐字符数对比），图片数量一一对应，水印 0 残留。

## License

MIT
