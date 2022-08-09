import os
import sys
import collections 
import collections.abc
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Cm

from functions.inputs import getDesiredInputs
from functions.readCSVFile import readCSVFile
from functions.slideEdits import editShape

def main():
  input("Before beginning please ensure all instances of powerpoint are closed and that your input.csv file is in the correct form (example: 21,Calibri,y,n ). Press Enter to begin... ") 

  try:
    os.mkdir('edited')
    print("NOTICE: edited directory created, continuing...")
  except OSError:
    print("NOTICE: edited directory already exists, continuing...")

  print("\nIMPORTANT: Please make sure that no instances of powerpoint are open before you begin!!!\n")

  textValuesDict = readCSVFile('input.csv')

  for filename in os.listdir("./"):

    if (filename[-4:] == ".ppt" or filename[-5:] == ".pptx"):

      try:
        prs = Presentation(filename)
      except BaseException as e:
        print("ERROR: " + str(e))
        print("WARNING: Failed to open powerpoint file: " + filename + " continuing...") 
        continue
      
      prs.slide_width = Cm(33.867)
      prs.slide_height = Cm(19.05)

      for slide in prs.slides:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(0, 0, 0)

        for shape in slide.shapes:
          editShape(filename, shape, textValuesDict)
    
      try:
        prs.save('edited/' + filename.split('.')[0] + '.pptx')
        print("Finished Successfully")
      except BaseException as error:
        print('ERROR STOPPING SCRIPT: Failed to save file: ' + '\'edited/' + filename.split('.')[0] + '.pptx\' ' + 'please make sure that all instances of powerpoint are closed') 
        sys.exit()

if __name__ == "__main__":
  main()