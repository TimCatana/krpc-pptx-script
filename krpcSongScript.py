import os
import sys
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE

try:
  os.mkdir('edited')
  print("edited directory created, continuing...")
except OSError as error:
  print("edited directory already exists, continuing...")

print("\nPlease make sure that no instances of powerpoint are open before you begin!!!\n")

for filename in os.listdir("./"):

  if filename != 'krpcSongScript.py' and filename != '.git' and filename != 'edited':
    
    # Font size
    fontSize = raw_input("enter desired font size: ")
    
    # Font Family
    fontFamily = raw_input("enter desired font (Calibri, etc..) if you input this correctly, the program will crash THIS IS CASE SENSITIVE. MOST FONTS CAPITALIZE THE FIRST LETTER: ")    
    
    # Font Bold
    fontBold = raw_input("bold? y or n: ")
    
    while (str(fontBold) != 'y' and str(fontBold) != 'n'):
      fontBold = raw_input("please enter a valid value - bold? y or n : ")
      
    if(str(fontBold) == 'y'):
      fontBold = True 
    else:
      fontBold = False

    # Font Italic
    fontItalic = raw_input("italic? y or n: ")

    while (str(fontItalic) != 'y' and str(fontItalic) != 'n'):
      fontItalic = raw_input("please enter a valid value - italic? y or n : ")

    if(str(fontItalic) == 'y'):
      fontItalic = True 
    else:
      fontItalic = False

    try:
      prs = Presentation(filename)
    except BaseException as error:
      print("Failed to open powerpoint file: " + filename) 

    for slide in prs.slides:
      slide.background.fill.solid()
      slide.background.fill.fore_color.rgb = RGBColor(0, 0, 0)

      for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
          if shape.has_text_frame:
            shape.text_frame.fit_text(font_family=str(fontFamily), max_size=int(fontSize), bold=fontBold, italic=fontItalic, font_file=None)
            for p in shape.text_frame.paragraphs:
              p.font.color.rgb = RGBColor(255, 255, 255)
        
        if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
          if shape.has_text_frame:
            shape.text_frame.fit_text(font_family=str(fontFamily), max_size=int(fontSize), bold=fontBold, italic=fontItalic, font_file=None)
            for p in shape.text_frame.paragraphs:
              p.font.color.rgb = RGBColor(255, 255, 255)

        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
          shape._element.getparent().remove(shape._element)

      try:
        prs.save('edited/' + filename.split('.')[0] + '.pptx')
      except BaseException as error:
        print('Failed to save file: ' + '\'edited/' + filename.split('.')[0] + '.pptx\' ' + 'please make sure that all instances of powerpoint are closed') 
        sys.exit()
