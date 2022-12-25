import collections 
import collections.abc

# Gets the desired font size from the user
# User is allowed to input empty string
# Default value is 20
def fontSizeInput():
  userInput = input("Enter desired font size: ")
    
  try:
    if(len(userInput) != 0):
      return int(userInput)
    else:
      print("NOTICE: You entered an empty string, using default font size of 20")
      return 20
  except Exception as e:
    print("ERROR: " + str(e))
    print("NOTICE: Something went wrong, using default font size of 20")
    return 20

# Gets the desired font family from the user
# User is allowed to input empty string
# Default value is 'Calibri'
def fontNameInput():
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

# Gets all the desired user inputs for the text format and puts them in a dictionary
# returns a dictionary with the desired valued for the text format
#
def getDesiredInputs():
  textValuesDict = {}

  textValuesDict['fontSize'] = fontSizeInput()
  textValuesDict['fontName'] = fontNameInput()
  textValuesDict['fontBold'] = yesOrNoInput("bold? y or n: ")
  textValuesDict['fontItalic'] = yesOrNoInput("italic? y or n: ")

  return textValuesDict
