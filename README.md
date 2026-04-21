# pdf-split
This is  a pdf split tool, supporting split automaticlly and by-hand

# PDF 智能拆分工具说明

双模式 PDF 拆分工具，通过选项卡切换。

## 依赖安装

```bash
pip install PyMuPDF Pillow opencv-python numpy paddlepaddle paddleocr
```

## 模式一：OCR 智能切分

对扫描件 PDF 进行自动拆分。工具使用 PaddleOCR 识别每页顶部 25% 区域的文字，当命中用户预设的关键字时执行切分。

**操作步骤：** 加载 PDF → 设置输出目录 → 输入关键字（每行一个）→ 点击启动

**分组命名规则：** 输出文件按 `线路N-关键字.pdf` 命名。当某个关键字在同一组内重复出现时，线路编号自动递增，实现按组归类。

## 模式二：手动预览切割

左侧预览 PDF 页面，右侧手动指定页码范围进行切割。

**操作步骤：** 加载 PDF → 翻页确认内容 → 填写起始/结束页码和文件名 → 点击切割

**辅助特性：**
- 翻页时自动填充结束页码
- 切割完成后起始页码自动跳到下一页，便于连续作业
- 支持缩放（0.4x–3.0x）和鼠标滚轮滚动
- 同名文件自动追加序号，不会覆盖
