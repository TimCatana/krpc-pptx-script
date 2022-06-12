import os
import sys
import collections 
import collections.abc
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE

# Gets the desired font size from the user
# User is allowed to input empty string
# Default value is 20
def fontSizeInput():
  userInput = input("Enter desired font size: ")
    
  try:
    if(len(userInput) != 0):
      return int(userInput)
    else:
      print("NOTICE: You entered an empty string, using default font family of 'Calibri'")
      return 20
  except Exception as e:
    print("ERROR: " + str(e))
    print("NOTICE: Something went wrong, using default font size of 20")
    return 20
  
# Gets the desired font family from the user
# User is allowed to input empty string
# Default value is 'Calibri'
def fontFamilyInput():
  userInput = input("Enter desired font family (Case sensitive e.g. 'Calibri' not 'calibri'): ")

  try:
    if(len(userInput) != 0):
      return userInput
    else:
      print("NOTICE: You entered an empty string, using default font family of 'Calibri'")
      return "Calibri"
  except Exception as e:
    print("ERROR: " + str(e))
    print("NOTICE: Something went wrong, using default font family of 'Calibri'")
    return "Calibri"

# Asks the user a yes and no question and gets a 'y' or 'n' answer
# User is asked repeatedly to enter a valid value until they enter either 'y', or 'n'
# Default value is 'n'
def yesOrNoInput(message):
  userInput = input(message)

  while (userInput.lower() != 'y' and userInput.lower() != 'n'):
    userInput = input("please enter a valid value - " + message)

  if(userInput.lower() == 'y'):
    return True
  elif(userInput.lower() == 'n'):
    return False
  else:
    print("NOTICE: Something went wrong, using default value of 'n'")
    return False

#
#
#
def editText(filename, shape, fontFamily, fontSize, fontBold, fontItalic):
  try:
    shape.text_frame.fit_text(font_family=fontFamily)
  except Exception as e:
    print("ERROR " + filename + ": " + str(e))
    print("NOTICE: " + filename + ":  something went wrong, using default font value of Calibri")
    shape.text_frame.fit_text(font_family="Calibri")
      
  try:
    shape.text_frame.fit_text(max_size=fontSize, bold=fontBold, italic=fontItalic)
  except Exception as e:
    print("ERROR " + filename + ": " + str(e))
    print("NOTICE: " + filename + ":  something went wrong, using default font values of fontSize=20, bold=False, italic=False")
    shape.text_frame.fit_text(max_size=20, bold=False, italic=False)
      
  for p in shape.text_frame.paragraphs:
    p.font.color.rgb = RGBColor(255, 255, 255)

#
#
#
def editShape(filename, shape, fontFamily, fontSize, fontBold, fontItalic):
  if shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
    if shape.has_text_frame:
      editText(filename, shape, fontFamily, fontSize, fontBold, fontItalic)
          
  if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
    if shape.has_text_frame:
      editText(filename, shape, fontFamily, fontSize, fontBold, fontItalic)

  if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
    shape._element.getparent().remove(shape._element)

#
#
#
def main():
  try:
    os.mkdir('edited')
    print("NOTICE: edited directory created, continuing...")
  except OSError:
    print("NOTICE: edited directory already exists, continuing...")

  print("\nIMPORTANT: Please make sure that no instances of powerpoint are open before you begin!!!\n")

  for filename in os.listdir("./"):

    if (filename[-4:] == ".ppt" or filename[-5:] == ".pptx"):
      fontSize = fontSizeInput()
      fontFamily = fontFamilyInput()
      fontBold = yesOrNoInput("bold? y or n: ")
      fontItalic = yesOrNoInput("italic? y or n: ")

      try:
        prs = Presentation(filename)
      except BaseException as e:
        print("ERROR: " + str(e))
        print("WARNING: Failed to open powerpoint file: " + filename + " continuing...") 
        continue
      
      for slide in prs.slides:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(0, 0, 0)

        for shape in slide.shapes:
          editShape(filename, shape, fontFamily, fontSize, fontBold, fontItalic)
    
      try:
        prs.save('edited/' + filename.split('.')[0] + '.pptx')
      except BaseException as error:
        print('ERROR STOPPING SCRIPT: Failed to save file: ' + '\'edited/' + filename.split('.')[0] + '.pptx\' ' + 'please make sure that all instances of powerpoint are closed') 
        sys.exit()

if __name__ == "__main__":
  main()