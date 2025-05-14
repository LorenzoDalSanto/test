from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE

def add_competitor_row(slide, data, y_offset):

    def add_textbox(text, left, top, width, height, font_size, font_name, bold, italic, color_rgb):
        textbox = slide.shapes.add_textbox(left, top, width, height)
        tf = textbox.text_frame

        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.word_wrap = True

        tf.clear()
        
        p = tf.paragraphs[0]
        
        run = p.add_run()
        run.text = text
        run.font.size = Pt(font_size)
        run.font.name = font_name
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = RGBColor(*color_rgb)

    # SOCIETÀ
    add_textbox(
        text=data["societa"],
        left=435600, top=1360800+y_offset, width=1076400, height=252000,
        font_size=10.5, font_name="Calibri", bold=False, italic=False, color_rgb=(0, 0, 0)
    )

    # SEDE
    add_textbox(
        text=data["sede"],
        left=3078000, top=1270800+y_offset , width=1076400, height=410400,
        font_size=10.5, font_name="Calibri", bold=False, italic=False, color_rgb=(0, 0, 0)
    )

    # DESCRIZIONE
    add_textbox(
        text=data["descrizione"],
        left=4233600, top=1191600+y_offset, width=3528000, height=252000,
        font_size=10.5, font_name="Calibri", bold=False, italic=False, color_rgb=(0, 0, 0)
    )

    # VDP
    add_textbox(
        text="VDP\n"+data["vdp"],
        left=7920000, top=1310400+y_offset, width=666000, height=363600,
        font_size=9, font_name="Calibri", bold=False, italic=False, color_rgb=(0, 0, 0)
    )

    # EBITDA
    add_textbox(
        text="EBITDA\n"+data["ebitda"],
        left=8575200, top=1310400+y_offset, width=666000, height=363600,
        font_size=9, font_name="Calibri", bold=False, italic=False, color_rgb=(0, 0, 0)
    )

    # EBITDA %
    add_textbox(
        text="EBITDA %\n"+data["ebitda_percent"],
        left=9230400, top=1310400+y_offset, width=666000, height=363600,
        font_size=9, font_name="Calibri", bold=False, italic=False, color_rgb=(0, 0, 0)
    )

    # N DIP
    add_textbox(
        text="N. Dip.\n"+data["n_dip"],
        left=9885600, top=1310400+y_offset, width=666000, height=363600,
        font_size=9, font_name="Calibri", bold=False, italic=False, color_rgb=(0, 0, 0)
    )
    # PFN
    add_textbox(
        text="PFN\n"+data["pfn"],
        left=10540800, top=1310400+y_offset, width=666000, height=363600,
        font_size=9, font_name="Calibri", bold=False, italic=False, color_rgb=(0, 0, 0)
    )

    # FCF
    add_textbox(
        text="FCF\n"+data["fcf"],
        left=11196000, top=1310400+y_offset, width=666000, height=363600,
        font_size=9, font_name="Calibri", bold=False, italic=False, color_rgb=(0, 0, 0)
    )
