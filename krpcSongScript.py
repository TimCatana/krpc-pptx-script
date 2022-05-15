import os
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE

for filename in os.listdir("./"):
  if filename != 'krpcSongScript.py':
    prs = Presentation(filename)
    for slide in prs.slides:
      slide.background.fill.solid()
      slide.background.fill.fore_color.rgb = RGBColor(0, 0, 0)

      for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
          if shape.has_text_frame:
            print (shape.shape_type)
            shape.text_frame.fit_text(font_family='Calibri', max_size=30, bold=False, italic=False, font_file=None)
            for p in shape.text_frame.paragraphs:
              p.font.color.rgb = RGBColor(255, 0, 0)
          if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            shape.getparent().remove(shape._element)
        # else if shape.shape_type 
    # prs.save(filename.split('.')[0] + ' EDITED.pptx')
