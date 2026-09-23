#!/usr/bin/env python3
"""1:1 recreation of the user's real Windows Terminal capture (ref-20.png).

Measured geometry (1483x578):
  y=0..10   tab strip padding          RGB(46,46,46)   #2E2E2E
  y=11..50  active tab chip x=12..312  RGB(12,12,12)   #0C0C0C (merges with body)
  y>=50     body                       RGB(12,12,12)   #0C0C0C
  first text baseline area y=64..78, x=12
  text RGB(204,204,204)  #CCCCCC
  line spacing 22px, font Consolas-like 16px
  block cursor inline after prompt

Usage:
    python generate_console.py out.png --cmd "java t3 ..." --code T3.java
"""

from __future__ import annotations

import argparse
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

# ---- measured from ref-20.png ----
TAB_STRIP = (46, 46, 46)        # #2E2E2E
BODY_BG = (12, 12, 12)          # #0C0C0C
FG = (204, 204, 204)            # #CCCCCC
GLYPH = (204, 204, 204)
CURSOR = (204, 204, 204)

STRIP_TOP = 11                  # gray strip above active tab chip
TAB_BOTTOM = 50                 # where tab chrome ends / body continues
TAB_X0, TAB_X1 = 12, 312        # active tab chip horizontal span
PAD_X = 12
PAD_Y = 37                      # first text y≈64 = 50+14
LINE_H = 23
FONT_SIZE = 16
BTN_W = 44

DEFAULT_W = 1479
DEFAULT_H = 574

MONO_CANDIDATES = [
    "C:/Windows/Fonts/consola.ttf",
    "C:/Windows/Fonts/CascadiaMono.ttf",
    "C:/Windows/Fonts/lucon.ttf",
    "C:/Windows/Fonts/cour.ttf",
]
CJK_CANDIDATES = [
    "C:/Windows/Fonts/msyh.ttc",
    "C:/Windows/Fonts/simsun.ttc",
    "C:/Windows/Fonts/simhei.ttf",
]
UI_CANDIDATES = [
    "C:/Windows/Fonts/segoeui.ttf",
    "C:/Windows/Fonts/msyh.ttc",
]


TAB_CHROME_PATH = Path(__file__).resolve().parent.parent / 'assets' / 'tab-chrome.png'


def load_font(size: int, candidates: list[str]):
    for path in candidates:
        if Path(path).is_file():
            try:
                return ImageFont.truetype(path, size=size)
            except Exception:
                continue
    return ImageFont.load_default()


def _decode_best(raw: bytes) -> str:
    if not raw:
        return ""
    for enc in ("utf-8", "gbk", "cp936", "utf-8-sig", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def run_java(code_path: Path, cmd_args: str, timeout: int = 20):
    work = Path(tempfile.mkdtemp(prefix="java-console-"))
    try:
        src = work / code_path.name
        src.write_text(code_path.read_text(encoding="utf-8"), encoding="utf-8")
        c = subprocess.run(
            ["javac", "-J-Dfile.encoding=UTF-8", "-encoding", "UTF-8", str(src)],
            capture_output=True, timeout=timeout, cwd=work,
        )
        if c.returncode != 0:
            return cmd_args, f"[编译失败]\n{_decode_best(c.stderr)}", False
        class_name = code_path.stem
        parts = cmd_args.split()
        if parts and parts[0] == "java":
            parts = parts[1:]
        if parts and parts[0].endswith(class_name):
            parts = parts[1:]
        r = subprocess.run(
            ["java", "-Dfile.encoding=UTF-8", "-Dsun.stdout.encoding=UTF-8",
             "-Dsun.stderr.encoding=UTF-8", class_name, *parts],
            capture_output=True, timeout=timeout, cwd=work,
        )
        raw = r.stdout if r.stdout else r.stderr
        out = _decode_best(raw).replace("�", "")
        return cmd_args, out.rstrip() if out.strip() else "(无输出)", r.returncode == 0
    except Exception as exc:  # noqa: BLE001
        return cmd_args, f"[执行异常] {exc}", False
    finally:
        try:
            for p in work.glob("*"):
                p.unlink(missing_ok=True)
            work.rmdir()
        except Exception:
            pass


def wrap_by_cells(text: str, max_cells: int = 110) -> list[str]:
    lines: list[str] = []
    for raw in (text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        n, buf = 0, []
        for ch in raw:
            w = 2 if ord(ch) > 0x2E80 else 1
            if n + w > max_cells:
                lines.append("".join(buf))
                buf, n = [], 0
            buf.append(ch)
            n += w
        lines.append("".join(buf))
    return lines or [""]


def draw_grid_line(draw, x0, y, text, font_ascii, font_cjk, cell_w, fill) -> float:
    """Draw one horizontal line on a shared baseline (水平对齐).

    ASCII = 1 cell, CJK = 2 cells. Mixed fonts share the same baseline y so
    the line does not stagger up/down.
    """
    x = float(x0)
    baseline_y = float(y)

    def _metrics(font):
        try:
            return font.getmetrics()  # (ascent, descent)
        except Exception:
            size = getattr(font, "size", 16)
            return int(size * 0.8), int(size * 0.2)

    asc_a, _ = _metrics(font_ascii)
    asc_c, _ = _metrics(font_cjk)

    for ch in text:
        if ord(ch) > 0x2E80:
            font, adv, asc = font_cjk, cell_w * 2, asc_c
        else:
            font, adv, asc = font_ascii, cell_w, asc_a
        # Pillow text y is top of em box; put baseline at baseline_y
        draw.text((x, baseline_y - asc), ch, fill=fill, font=font)
        x += adv
    return x


def generate_console_image(
    out_path: Path,
    command: str,
    output: str,
    *,
    title: str = r"C:\Windows\system32\cmd.exe",
    prompt: str = r"C:\Users\XiaoMibook>",
    width: int = DEFAULT_W,
    height: int = DEFAULT_H,
    font_size: int = FONT_SIZE,
) -> Path:
    ss = 2
    # Consolas 16px advance ≈ 8.8px — use 9 for grid
    cell_w = 9

    # java-help capture style: no banner on runs; empty session keeps banner
    content_lines = []
    if not command and not output:
        content_lines += [
            "Microsoft Windows [版本 10.0.26200.9457]",
            "(c) Microsoft Corporation。保留所有权利。",
            "",
            prompt,
        ]
    else:
        content_lines.append(f"{prompt}{command}" if command else prompt)
        if output:
            content_lines.extend(wrap_by_cells(output, max_cells=200))
        content_lines.append("")
        content_lines.append(prompt)

    # keep 1:1 canvas size (or grow if lots of output)
    body_need = PAD_Y + (len(content_lines) + 1) * LINE_H + 8
    win_h = max(height, TAB_BOTTOM + body_need)
    win_w = width

    ascii_font = load_font(font_size * ss, MONO_CANDIDATES)
    cjk_font = load_font(font_size * ss, CJK_CANDIDATES)
    ui_font = load_font(12 * ss, UI_CANDIDATES)
    cell_s = cell_w * ss
    line_s = LINE_H * ss

    canvas = Image.new("RGB", (win_w * ss, win_h * ss), BODY_BG)
    draw = ImageDraw.Draw(canvas)

    def S(v):
        return int(round(v * ss))

    # full body
    draw.rectangle([0, 0, S(win_w) - 1, S(win_h) - 1], fill=BODY_BG)

    # ---- 1:1 paste the user's tab-chrome bitmap (do NOT redraw) ----
    if TAB_CHROME_PATH.is_file():
        chrome = Image.open(TAB_CHROME_PATH).convert("RGB")
        cw, ch = chrome.size  # 373 x 47
        # scale by ss for the supersampled canvas
        chrome_ss = chrome.resize((cw * ss, ch * ss), Image.Resampling.NEAREST)
        canvas.paste(chrome_ss, (0, 0))
        # extend tab strip to full width using measured colors
        # gray top strip y=0..3, right panel x>=301
        draw.rectangle([S(cw), 0, S(win_w) - 1, S(4)], fill=TAB_STRIP)
        draw.rectangle([S(301), S(4), S(win_w) - 1, S(43)], fill=TAB_STRIP)
        # body under chrome continues (already BODY_BG)
        draw.rectangle([0, S(44), S(win_w) - 1, S(win_h) - 1], fill=BODY_BG)
    else:
        # fallback if asset missing
        draw.rectangle([0, 0, S(win_w) - 1, S(STRIP_TOP) - 1], fill=TAB_STRIP)
        draw.rectangle([S(TAB_X1), S(STRIP_TOP), S(win_w) - 1, S(TAB_BOTTOM) - 1], fill=TAB_STRIP)
        draw.rectangle([0, S(STRIP_TOP), S(TAB_X0) - 1, S(TAB_BOTTOM) - 1], fill=TAB_STRIP)
        draw.rounded_rectangle(
            [S(TAB_X0), S(STRIP_TOP), S(TAB_X1), S(TAB_BOTTOM) + S(20)],
            radius=S(6), fill=BODY_BG,
        )
        draw.rectangle([0, S(TAB_BOTTOM), S(win_w) - 1, S(win_h) - 1], fill=BODY_BG)

    # window controls only on the far right of the strip (not in the 373px crop)
    for i, kind in enumerate(("min", "max", "close")):
        x0 = S(win_w) - S((3 - i) * BTN_W)
        x1 = x0 + S(BTN_W)
        cx = (x0 + x1) // 2
        cyb = S(22)
        lw = max(1, ss // 1)
        if kind == "min":
            draw.line([cx - S(7), cyb, cx + S(7), cyb], fill=GLYPH, width=lw)
        elif kind == "max":
            draw.rectangle([cx - S(6), cyb - S(6), cx + S(6), cyb + S(6)], outline=GLYPH, width=lw)
        else:
            draw.line([cx - S(5), cyb - S(5), cx + S(5), cyb + S(5)], fill=GLYPH, width=lw)
            draw.line([cx + S(5), cyb - S(5), cx - S(5), cyb + S(5)], fill=GLYPH, width=lw)

    # --- body text starting y=64 ---
    y = S(TAB_BOTTOM + PAD_Y)
    last_x = S(PAD_X)
    prompt_line_idx = len(content_lines) - 1
    cursor_at = None
    for i, line in enumerate(content_lines):
        end_x = draw_grid_line(draw, S(PAD_X), y, line, ascii_font, cjk_font, cell_s, FG)
        if line:
            last_x = end_x
        if i == prompt_line_idx:
            cursor_at = (end_x, y)
        y += line_s

    # inline block cursor after prompt (not on a new line)
    if cursor_at:
        cx0, cy0 = cursor_at
        ch_ = int(FONT_SIZE * ss * 0.9)
        draw.rectangle(
            [cx0 + S(1), cy0 + S(1), cx0 + cell_s - S(1), cy0 + S(1) + ch_],
            fill=CURSOR,
        )

    # thin scrollbar on the right (in user's real capture)
    sb_x0 = S(win_w - 8)
    sb_x1 = S(win_w - 1)
    draw.rectangle([sb_x0, S(48), sb_x1, S(win_h) - 1], fill=(28, 28, 28))
    draw.rectangle([sb_x0 + S(2), S(70), sb_x1 - S(2), S(win_h) - S(40)], fill=(144, 144, 144))

    out = canvas.resize((win_w, win_h), Image.Resampling.LANCZOS)

    # 1:1 paste user's tab chrome AFTER downscale (pixel-exact)
    if TAB_CHROME_PATH.is_file():
        chrome = Image.open(TAB_CHROME_PATH).convert('RGB')
        out.paste(chrome, (0, 0))

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out.save(out_path, format='PNG')
    return out_path


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="1:1 Windows Terminal screenshot (ref-20)")
    parser.add_argument("out", type=Path)
    parser.add_argument("--cmd", default="")
    parser.add_argument("--output-text", default="")
    parser.add_argument("--code", type=Path, default=None)
    parser.add_argument("--prompt", default=r"C:\Users\XiaoMibook>")
    parser.add_argument("--title", default=r"C:\Windows\system32\cmd.exe")
    parser.add_argument("--font-size", type=int, default=FONT_SIZE)
    parser.add_argument("--width", type=int, default=DEFAULT_W)
    parser.add_argument("--height", type=int, default=DEFAULT_H)
    args = parser.parse_args(argv[1:])

    command = args.cmd
    output = args.output_text
    if args.code and args.code.is_file():
        run_cmd = command or f"java {args.code.stem}"
        command, output, _ok = run_java(args.code, run_cmd)

    generate_console_image(
        args.out, command, output,
        title=args.title, prompt=args.prompt,
        font_size=args.font_size, width=args.width, height=args.height,
    )
    print(f"OK: {args.out.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
