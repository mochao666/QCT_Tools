# -*- coding: utf-8 -*-
"""生成 QCT 小工具图标：放大镜 + 对勾，深蓝 + 浅灰。运行后得到 icon.ico。"""
import os
try:
    from PIL import Image, ImageDraw
except ImportError:
    print("请先安装 Pillow: pip install Pillow")
    raise

# 主色：深蓝（精准、专业）、浅灰（背景）
DARK_BLUE = "#1a365d"
LIGHT_GRAY = "#e8e8e8"
WHITE = "#ffffff"

def draw_icon(size):
    """绘制单尺寸图标：浅灰底 + 放大镜 + 对勾。"""
    img = Image.new("RGBA", (size, size), LIGHT_GRAY)
    d = ImageDraw.Draw(img)
    cx, cy = size // 2, size // 2
    r_lens = size // 4
    lens_cx = cx - size // 10
    lens_cy = cy - size // 10
    # 放大镜镜片（圆），加粗描边更清晰
    d.ellipse(
        [lens_cx - r_lens, lens_cy - r_lens, lens_cx + r_lens, lens_cy + r_lens],
        outline=DARK_BLUE,
        width=max(3, size // 24),
        fill=WHITE,
    )
    # 放大镜柄，加粗
    handle_w = max(3, size // 16)
    x1 = lens_cx + int(r_lens * 0.7)
    y1 = lens_cy + int(r_lens * 0.7)
    x2 = cx + size // 3
    y2 = cy + size // 3
    d.line([(x1, y1), (x2, y2)], fill=DARK_BLUE, width=handle_w)
    # 对勾（在镜片内）
    x0 = lens_cx - r_lens // 2
    y0 = lens_cy
    x1c = lens_cx - r_lens // 4
    y1c = lens_cy + r_lens // 2
    x2c = lens_cx + r_lens
    y2c = lens_cy - r_lens // 2
    d.line([(x0, y0), (x1c, y1c), (x2c, y2c)], fill=DARK_BLUE, width=max(2, size // 36))
    return img

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    sizes = [256, 48, 32, 16]
    images = [draw_icon(s) for s in sizes]
    out_path = os.path.join(script_dir, "icon.ico")
    images[0].save(
        out_path,
        format="ICO",
        sizes=[(s, s) for s in sizes],
        append_images=images[1:],
    )
    print("已生成:", out_path)

if __name__ == "__main__":
    main()
