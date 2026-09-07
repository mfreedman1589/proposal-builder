"""report_charts.py -- chart/map PNGs for the attribution report deck.

Pillow-only rendering -- no Streamlit, no DB, and deliberately no new
pinned dependency (matplotlib isn't in requirements.txt and this project's
existing chart/map feature, targeting_map.py, already established "Pillow
alone, no filesystem font lookup" as the house convention for exactly this
kind of rendering -- see CLAUDE.md's Zip/map builder note). Every image is
rendered at the real EMU footprint of its named template region
(ChartRegion/MapRegion), converted to inches and then to pixels at
PRINT_DPI, never a screen-preview default -- because these slides get
copied into proposal decks later (create-case-study-from-report, the slide
vault) and have to hold up printed. This is also why ATTRIBUTION_REPORT_
PLAN.md rules out native pptx charts for this deck: an image survives a
cross-deck slide copy the way a live pptx chart object never does (see
CLAUDE.md's case-study-copying rules -- "a chart is not a leaf").
"""
import io

from PIL import Image, ImageDraw, ImageFont

import targeting_map

PRINT_DPI = 200
_EMU_PER_INCH = 914400
_BAR_COLOR = (0, 9, 70, 255)       # navy 000946, this deck's own header color


def region_pixels(width_emu, height_emu, dpi=PRINT_DPI):
    """(width_px, height_px) for a named template region at print DPI."""
    return (max(1, round(width_emu / _EMU_PER_INCH * dpi)),
            max(1, round(height_emu / _EMU_PER_INCH * dpi)))


def _font(size):
    return ImageFont.load_default(size=size)


def render_bar_chart(labels, values, width_emu, height_emu, value_labels=None,
                     color=_BAR_COLOR):
    """A horizontal bar chart PNG (bytes), one bar per (label, value) pair,
    sized to a named ChartRegion's own EMU footprint. Draws exactly the
    rows it's given -- no ranking, filtering or truncation of its own; the
    caller (report_assembly) has already picked what belongs on the chart.
    `value_labels`, if given, is one already-formatted string per bar (e.g.
    "1.36%") drawn at the end of its bar -- never recomputed here from the
    raw value, so the chart can never disagree with the table showing the
    same rows. None when there's nothing to plot, matching
    assembly.place_targeting_map's own "nothing to draw" convention.
    """
    if not labels or not values:
        return None
    width_px, height_px = region_pixels(width_emu, height_emu)
    img = Image.new("RGBA", (width_px, height_px), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    n = len(labels)
    font_size = max(10, min(16, int(height_px / n * 0.32)))
    font = _font(font_size)
    row_h = height_px / n
    bar_h = row_h * 0.5
    label_w = max(draw.textlength(str(label), font=font) for label in labels) + 14
    max_value = max(values) or 1
    value_label_w = (max(draw.textlength(str(v), font=font) for v in value_labels) + 10
                     if value_labels else 0)
    bar_area_w = max(10, width_px - label_w - value_label_w - 6)
    for i, (label, value) in enumerate(zip(labels, values)):
        y_center = row_h * i + row_h / 2
        top, bottom = y_center - bar_h / 2, y_center + bar_h / 2
        draw.text((0, y_center), str(label), font=font, fill=color, anchor="lm")
        bar_w = max(2, bar_area_w * (value / max_value)) if value else 2
        left = label_w
        draw.rectangle([left, top, left + bar_w, bottom], fill=color)
        if value_labels:
            draw.text((left + bar_w + 6, y_center), str(value_labels[i]), font=font,
                      fill=color, anchor="lm")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def render_zip_map(zip_values, width_emu, height_emu, legend_title="Attributed rate"):
    """The zip slide's choropleth -- zips shaded by attributed rate,
    quantile-binned, via targeting_map.render_choropleth.

    `zip_values` is {zip: attributed_rate}. Returns `(png, missing_zips)` --
    zips with a value but no ZCTA polygon, which the caller reports rather
    than dropping silently. `(None, [])` when there is nothing to plot,
    matching assembly.place_targeting_map's own convention, so the caller
    leaves the template's placeholder region alone.

    Requires `market_lookup.install()` for place labels to resolve;
    degrades to a plainer map rather than raising when it hasn't been
    called.
    """
    if not zip_values:
        return None, []
    width_px, height_px = region_pixels(width_emu, height_emu)
    return targeting_map.render_choropleth(
        zip_values, width_px=width_px, height_px=height_px, legend_title=legend_title)
