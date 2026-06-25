# PPT 导出 + AI 视觉审查工具

把 PowerPoint 幻灯片导出为高清图片，配合 AI 视觉模型自动检查排版质量。

## 能查出什么问题？

| 检查项 | 说明 |
|--------|------|
| 🔴 元素碰撞 | 箭头戳进卡片、文本框盖住图片 |
| 🔴 文字溢出 | 文本框太小，字跑到框外面 |
| 🟡 边距不足 | 内容贴着幻灯片边缘，投影时可能被裁 |
| 🟡 对比度过低 | 灰色字在深色背景上，后排看不清 |
| 🟡 间距不均 | 左边空一大片，右边挤成一团 |
| 🟢 占位符残留 | 模板里忘删的假文字 |

## 环境要求

- **Windows** 系统
- **Microsoft PowerPoint** 已安装（Office 2016+ 或 Microsoft 365）
- **Python 3.7+** + `pywin32` 包

```bash
pip install pywin32
```

## 使用方法

### 第一步：导出图片

```bash
python export_slides.py <你的PPT文件.pptx>
```

示例：
```bash
python export_slides.py "C:\Users\a\Desktop\我的汇报.pptx"
```

执行后会在 PPT 同目录下生成 `slides_img/` 文件夹，里面是逐页的高清 PNG 截图。

### 第二步：AI 视觉审查

把导出的图片交给任意支持视觉的 AI（ChatGPT、Claude、Gemini 等），用这段提示词：

```
请逐张检查这些 PPT 幻灯片图片，重点关注：
1. 元素是否重叠、碰撞
2. 文字是否溢出文本框
3. 边距是否足够（至少 0.5 英寸）
4. 文字/图标对比度是否足够
5. 间距是否均匀
6. 是否有残留的占位符文字

对每页逐一报告问题。
```

### 第三步：修复问题

根据 AI 的报告逐个修复 PPT 中的排版问题。

---

## 工作原理

```
你的 PPT
   ↓  export_slides.py（操控 PowerPoint 逐页导出）
PNG 高清图片
   ↓  喂给 AI 视觉模型
排版问题报告
   ↓  你在 PowerPoint 里手工修
完美 PPT ✅
```

---

## 常见问题

**Q: 报错 `ModuleNotFoundError: No module named 'win32com'`？**
```bash
pip install pywin32
```

**Q: PowerPoint 弹窗关不掉？**
脚本会自动关闭。如果卡住了，在任务管理器里结束 POWERPNT.EXE。

**Q: 能在 macOS 上用吗？**
不行。macOS 版本需要用 LibreOffice 替代方案（`soffice --headless --convert-to png`）。
