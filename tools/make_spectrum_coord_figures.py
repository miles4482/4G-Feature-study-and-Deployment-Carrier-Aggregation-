#!/usr/bin/env python3
"""Schematics for the LTE Spectrum Coordination sheet (not Huawei originals)."""
import os
from PIL import Image, ImageDraw, ImageFont

OUT = "/workspace/ca_figures"
NAVY = (31, 78, 121)
NAVY2 = (46, 117, 182)
TEAL = (13, 115, 119)
GOLD = (201, 162, 39)
GREEN = (84, 130, 53)
ORANGE = (198, 89, 17)
RED = (192, 0, 0)
GRAY = (89, 89, 89)
LIGHT = (242, 242, 242)
WHITE = (255, 255, 255)
PALE_BLUE = (214, 234, 248)
PALE_GOLD = (255, 242, 204)
PALE_GREEN = (226, 239, 218)
PALE_ORANGE = (252, 228, 214)


def _font(size, bold=False):
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    ]
    for p in candidates:
        if os.path.exists(p):
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()


def _round_rect(d, xy, r, fill, outline=None, width=1):
    d.rounded_rectangle(xy, radius=r, fill=fill, outline=outline, width=width)


def _center(d, box, text, font, fill):
    x0, y0, x1, y1 = box
    bbox = d.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    d.text((x0 + (x1 - x0 - tw) / 2, y0 + (y1 - y0 - th) / 2), text, font=font, fill=fill)


def make_ul_dl_decoupling():
    w, h = 1400, 780
    im = Image.new("RGB", (w, h), WHITE)
    d = ImageDraw.Draw(im)
    title_f = _font(28, True)
    h_f = _font(18, True)
    b_f = _font(15)
    s_f = _font(13)
    d.text((40, 18), "LTE Spectrum Coordination  —  UL/DL channel decoupling on a CA UE", font=title_f, fill=NAVY)
    d.text((40, 56), "DL stays on both high and low bands. UL is selected: cell-center → high band, cell-edge → low band.", font=b_f, fill=GRAY)

    # two panels
    panels = [
        (40, 100, 680, 740, "Cell-center UE", GREEN, PALE_GREEN,
         "UL on HIGH band (capacity)",
         "DL on HIGH + LOW (CA)",
         "PCell = high-band cell\nSCell = low-band cell\nUL PUSCH / PUCCH on PCell (high)\nDL data on both CCs"),
        (720, 100, 1360, 740, "Cell-edge UE", ORANGE, PALE_ORANGE,
         "UL on LOW band (coverage)",
         "DL on HIGH + LOW (CA)",
         "PCell = low-band cell\nSCell = high-band cell\nUL PUSCH / PUCCH on PCell (low)\nDL data on both CCs  →  cell-edge DL rate gain"),
    ]
    for x0, y0, x1, y1, title, accent, pale, ul, dl, note in panels:
        _round_rect(d, (x0, y0, x1, y1), 16, pale, outline=accent, width=3)
        _round_rect(d, (x0, y0, x1, y0 + 48), 16, accent)
        d.rectangle((x0, y0 + 24, x1, y0 + 48), fill=accent)
        _center(d, (x0, y0, x1, y0 + 48), title, h_f, WHITE)

        # eNodeB
        _round_rect(d, (x0 + 180, y0 + 70, x1 - 180, y0 + 130), 10, NAVY)
        _center(d, (x0 + 180, y0 + 70, x1 - 180, y0 + 130), "eNodeB  (same site, overlapping coverage)", s_f, WHITE)

        # two bands
        hx0, hx1 = x0 + 40, x0 + 290
        lx0, lx1 = x1 - 290, x1 - 40
        by0, by1 = y0 + 160, y0 + 250
        _round_rect(d, (hx0, by0, hx1, by1), 10, (192, 80, 77), outline=RED, width=2)
        _center(d, (hx0, by0, hx1, by1), "HIGH band cell\n(e.g. 1800 / TDD mid)", s_f, WHITE)
        _round_rect(d, (lx0, by0, lx1, by1), 10, NAVY2, outline=NAVY, width=2)
        _center(d, (lx0, by0, lx1, by1), "LOW band cell\n(e.g. 800 / 900)", s_f, WHITE)

        # UE
        ux0, uy0 = x0 + 200, y0 + 430
        ux1, uy1 = x1 - 200, y0 + 500
        _round_rect(d, (ux0, uy0, ux1, uy1), 12, GOLD)
        _center(d, (ux0, uy0, ux1, uy1), "CA UE  (Rel-10+)", h_f, WHITE)

        # arrows DL (both bands, dashed-like double)
        mid_h = ((hx0 + hx1) / 2, by1)
        mid_l = ((lx0 + lx1) / 2, by1)
        umid = ((ux0 + ux1) / 2, uy0)
        d.line([mid_h, (umid[0] - 40, umid[1])], fill=NAVY2, width=4)
        d.line([mid_l, (umid[0] + 40, umid[1])], fill=NAVY2, width=4)
        d.text((x0 + 50, y0 + 330), "DL  (always both bands)", font=s_f, fill=NAVY)

        # UL arrow — one band only
        if "HIGH" in ul:
            d.line([(umid[0] - 40, uy0), mid_h], fill=RED, width=6)
            d.polygon([(mid_h[0] - 10, by1 + 14), (mid_h[0] + 10, by1 + 14), (mid_h[0], by1 + 2)], fill=RED)
            d.text((hx0 + 10, y0 + 360), "UL  →  high band", font=h_f, fill=RED)
        else:
            d.line([(umid[0] + 40, uy0), mid_l], fill=NAVY, width=6)
            d.polygon([(mid_l[0] - 10, by1 + 14), (mid_l[0] + 10, by1 + 14), (mid_l[0], by1 + 2)], fill=NAVY)
            d.text((lx0 - 20, y0 + 360), "UL  →  low band", font=h_f, fill=NAVY)

        # captions
        _round_rect(d, (x0 + 24, y0 + 520, x1 - 24, y0 + 570), 8, WHITE, outline=accent, width=2)
        d.text((x0 + 40, y0 + 532), ul, font=h_f, fill=accent)
        d.text((x0 + 40, y0 + 552), dl, font=b_f, fill=NAVY)
        d.multiline_text((x0 + 40, y0 + 585), note, font=s_f, fill=GRAY, spacing=4)

    path = os.path.join(OUT, "spectrum_coord_ul_dl.png")
    im.save(path, "PNG")
    return path


def make_pcell_flow():
    w, h = 1400, 520
    im = Image.new("RGB", (w, h), WHITE)
    d = ImageDraw.Draw(im)
    title_f = _font(26, True)
    h_f = _font(16, True)
    s_f = _font(13)
    d.text((40, 16), "PCell change based on uplink quality  (SpectrumCoordinationSwitch)", font=title_f, fill=NAVY)
    d.text((40, 52), "Huawei CaMgtCfg.CellCaAlgoSwitch description: with spectrum coordination enabled, PCell changes are triggered based on uplink quality.  FDD and TDD.", font=s_f, fill=GRAY)

    boxes = [
        (40, 110, 280, 250, NAVY, "1. CA already live\nDL 2CC (or more)\nHigh + low overlap\nSCell activated"),
        (320, 110, 560, 250, TEAL, "2. Switch ON\nSpectrumCoordination\nSwitch = 1\non participating cells"),
        (600, 110, 840, 250, GOLD, "3. Monitor UL quality\non current PCell\n(PUSCH / UL SINR\nfamily — see FPD)"),
        (880, 110, 1120, 250, ORANGE, "4. If UL poor on\nhigh-band PCell\n(typical cell edge)\n→ PCell → low band"),
        (1160, 110, 1360, 250, GREEN, "5. Keep DL CA\nHigh band stays\nSCell for DL\nUL rides low band"),
    ]
    for i, (x0, y0, x1, y1, col, txt) in enumerate(boxes):
        _round_rect(d, (x0, y0, x1, y1), 12, col)
        d.multiline_text((x0 + 16, y0 + 18), txt, font=s_f, fill=WHITE, spacing=4)
        if i < len(boxes) - 1:
            d.polygon([(x1 + 4, (y0 + y1) / 2 - 10), (x1 + 18, (y0 + y1) / 2), (x1 + 4, (y0 + y1) / 2 + 10)], fill=NAVY2)

    # return path
    _round_rect(d, (40, 290, 1360, 480), 12, LIGHT, outline=NAVY2, width=2)
    d.text((60, 308), "Return / leave (cell-center again) and NSA exception", font=h_f, fill=NAVY)
    d.multiline_text(
        (60, 340),
        "• If UL quality on the high band recovers (UE moves toward cell center), PCell can change back to the high band so UL uses the capacity layer.\n"
        "• Turning SpectrumCoordinationSwitch Off stops UL-quality PCell changes; existing CA SCell management (A2 / traffic / CQI) continues as in CA Ch.4.6.\n"
        "• NSA DC: after Spectrum Coordination is enabled, ENBCELLRSVDPARA.RsvdSwPara6 bit20 can disable it for NSA DC UEs so those UEs are not handed over from PCC to a non-candidate PCC.\n"
        "• NSA DC PCC anchoring takes precedence over LTE spectrum coordination enhancement (WbbCaMultiCarrierCoordSw). If both are on, only NSA PCC anchoring takes effect.\n"
        "• This flow is reconstructed from the CellCaAlgoSwitch meaning + CloudAIR/Turkcell description. Dedicated LTE Spectrum Coordination FPD is not in this repository.",
        font=s_f, fill=GRAY, spacing=5,
    )
    path = os.path.join(OUT, "spectrum_coord_pcell_flow.png")
    im.save(path, "PNG")
    return path


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    print(make_ul_dl_decoupling())
    print(make_pcell_flow())
