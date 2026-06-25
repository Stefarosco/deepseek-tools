"""
PPT 幻灯片导出工具 — 将 .pptx 文件逐页导出为高清 PNG 图片
=================================================================
用途：搭配 AI 视觉模型进行 PPT 排版质量审查（QA）
原理：通过 Windows PowerPoint COM 接口操控 PPT 逐页导出
要求：Windows + 已安装 Microsoft PowerPoint

用法：
  python export_slides.py <PPT文件路径> [输出文件夹]

示例：
  python export_slides.py 我的汇报.pptx
  python export_slides.py C:\ppt\答辩.pptx D:\slides_img
"""

import os
import sys
import win32com.client


def export_slides(pptx_path, output_dir=None, width=1920, height=1080):
    """
    将 PPT 所有页面导出为 PNG 图片

    :param pptx_path:  PPT 文件路径
    :param output_dir: 图片输出目录（默认与 PPT 同目录下的 slides_img/）
    :param width:      导出宽度（像素），默认 1920
    :param height:     导出高度（像素），默认 1080
    :return:           导出的图片路径列表
    """
    if not os.path.exists(pptx_path):
        print(f"❌ 文件不存在: {pptx_path}")
        return []

    if output_dir is None:
        ppt_dir = os.path.dirname(os.path.abspath(pptx_path))
        ppt_name = os.path.splitext(os.path.basename(pptx_path))[0]
        output_dir = os.path.join(ppt_dir, "slides_img")

    os.makedirs(output_dir, exist_ok=True)

    print(f"📄 PPT: {pptx_path}")
    print(f"📁 输出: {output_dir}")
    print(f"📐 分辨率: {width}×{height}")
    print("-" * 50)

    exported = []
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True

    try:
        presentation = powerpoint.Presentations.Open(pptx_path, WithWindow=False)
        total = len(presentation.Slides)

        for i, slide in enumerate(presentation.Slides):
            img_name = f"slide-{i+1:02d}.png"
            img_path = os.path.join(output_dir, img_name)
            slide.Export(img_path, "PNG", width, height)
            exported.append(img_path)
            print(f"  ✅ [{i+1}/{total}] {img_name}")

        presentation.Close()
        print("-" * 50)
        print(f"🎉 完成！共导出 {total} 张图片 → {output_dir}")

    finally:
        powerpoint.Quit()

    return exported


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    pptx = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    export_slides(pptx, out)
