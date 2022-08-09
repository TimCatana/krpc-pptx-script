import collections 
import collections.abc
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Cm

# Takes the text and puts it on 2 lines 
#
# 
def editTextFormat(filename, shape):
  lines = shape.text.splitlines(False)

  for line in lines:
    if(line == '' or line.isspace() == True):
      lines.remove(line)
    # TODO - get rid of whitespace on the end of the strings, strip() was not working when I tried implementing it

  # Leave these here for future testing purposes
  # print(lines)
  # print(len(lines))

  addNewLine = len(lines) / 2 # This makes it so that all text is made to be 2 lines. See the for loop below
  editedText = ''

  for i in range(len(lines)):
    lines[i].rstrip()
    if(i == 0):
      editedText = lines[i] + ' '
    elif(i % addNewLine == 0):
      editedText = editedText + '\n' + lines[i] + ' '
    else:
      editedText = editedText + lines[i] + ' '

  shape.text = editedText

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
      editTextFormat(filename, shape)
      editTextStyle(filename, shape, textValuesDict)
          
  if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
    if shape.has_text_frame:
      shape.width = Cm(33.867)
      shape.height = Cm(19.05)
      shape.left = Cm(0)
      shape.top = Cm(0)
      editTextFormat(filename, shape)
      editTextStyle(filename, shape, textValuesDict)
      
  else :  
    shape._element.getparent().remove(shape._element)