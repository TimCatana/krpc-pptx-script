import collections 
import collections.abc
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Cm

# 
#
# TODO - theres probably a built in string function for this
def editTextFormat(filename, shape):
  newLines = 1
  for char in shape.text:
    print(repr(char))
    if(char == "\n" or char == "\x0b"):
      if(newLines % 2 == 0):
        print("edit char")
        newLines += 1
      else:
        newLines += 1
        print("dont edit char")
  

# Edits the text of the slide given the inputs
# 
#
def editTextStyle(filename, shape, textValuesDict):
  try:
    shape.text_frame.fit_text(font_family=textValuesDict['fontName'])
  except Exception as e:
    print("ERROR " + filename + ": " + str(e))
    print("NOTICE: " + filename + ":  something went wrong, using default font value of Calibri")
    shape.text_frame.fit_text(font_family="Calibri")
 
  try:
    shape.text_frame.fit_text(max_size=textValuesDict['fontSize'], bold=textValuesDict['fontBold'], italic=textValuesDict['fontItalic'])
  except Exception as e:
    print("ERROR " + filename + ": " + str(e))
    print("NOTICE: " + filename + ":  something went wrong, using default font values of fontSize=20, bold=False, italic=False")
    shape.text_frame.fit_text(max_size=20, bold=False, italic=False)
      
  for p in shape.text_frame.paragraphs:
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER

# Edits other shapes on the slide
# Edits text using desired values
# Gets rid of pictures
def editShape(filename, shape, textValuesDict):
  if shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
    if shape.has_text_frame:
      shape.width = Cm(33.867)
      shape.height = Cm(19.05)
      shape.left = Cm(0)
      shape.top = Cm(0)
      editTextStyle(filename, shape, textValuesDict)
      editTextFormat(filename, shape)
          
  if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
    if shape.has_text_frame:
      shape.width = Cm(33.867)
      shape.height = Cm(19.05)
      shape.left = Cm(0)
      shape.top = Cm(0)
      editTextStyle(filename, shape, textValuesDict)
      editTextFormat(filename, shape)
      
  else :  
    shape._element.getparent().remove(shape._element)