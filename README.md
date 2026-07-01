# DeepSeek 小工具集

各种提升效率的 Python 小工具，配合 AI 助手使用效果更佳。

## 工具列表

### 📊 Excel 可视化工具 (`excel_viz.py`)

让 AI 助手能"看见" Excel — 快速摸底、HTML 视觉预览、差异对比。

```bash
pip install openpyxl
python excel_viz.py dump  文件.xlsx           # 查看 Sheet 结构
python excel_viz.py html  文件.xlsx preview.html  # 导出带样式的 HTML
python excel_viz.py diff  旧.xlsx 新.xlsx diff.html  # 对比差异
```

详见下方 [excel_viz 详细说明](#excel_viz-使用说明)。

---

### 📽 PPT 幻灯片导出工具 (`export_slides.py`)

将 PPT 导出为高清 PNG，配合 AI 视觉模型审查排版质量。

```bash
pip install pywin32
python export_slides.py 我的汇报.pptx
```

执行后在 PPT 同目录生成 `slides_img/` 文件夹，逐页高清 PNG。
配合这段提示词让 AI 审查：

> 请逐张检查这些 PPT 幻灯片图片，重点关注：
> 1. 元素是否重叠、碰撞
> 2. 文字是否溢出文本框
> 3. 边距是否足够（至少 0.5 英寸）
> 4. 文字/图标对比度是否足够
> 5. 间距是否均匀
> 6. 是否有残留的占位符文字

**环境要求：** Windows + 已安装 Microsoft PowerPoint + `pywin32`

---

## excel_viz 使用说明

### 📋 dump — 快速摸底

```bash
python excel_viz.py dump D:\表格\销售数据.xlsx
```

输出所有 Sheet 的行列数、合并单元格、公式、数据预览。

### 🎨 html — 视觉预览

```bash
python excel_viz.py html D:\表格\销售数据.xlsx preview.html
```

生成的 HTML 保留颜色、边框、合并单元格，浏览器打开后截图喂给视觉 AI 做排版审查。

### 🔍 diff — 差异对比

```bash
python excel_viz.py diff 旧版.xlsx 新版.xlsx diff.html
```

终端输出所有不同单元格，可选生成高亮 HTML。

**环境要求：** Python 3.7+ + `openpyxl`（全平台通用）
