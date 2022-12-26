import csv

#
#
#


def readCSVFile(filename):
    textValuesDict = {}

    try:
        csvFile = open(filename)
    except:
        print("ERROR: you forgot the input.csv file")
        input("PRESS ENTER TO CLOSE TERMINAL")

    with csvFile:

        try:
            csvReader = csv.reader(csvFile, delimiter=',')
        except:
            print("ERROR: failed to parse csv file... that's all I know")
            input("PRESS ENTER TO CLOSE TERMINAL")

        for row in csvReader:

            try:
                row[0]
            except IndexError:
                print("ERROR: Font Size does not exist in input.csv")
                print(
                    "ERROR: input.csv should be in form 'fontSize,fontName,fontBold(y or n),fontItalic (y or n)")
                input("PRESS ENTER TO CLOSE TERMINAL")

            try:
                row[1]
            except IndexError:
                print("ERROR: Font Name does not exist in input.csv")
                print(
                    "ERROR: input.csv should be in form 'fontSize(number),fontName,fontBold(y or n),fontItalic (y or n)")
                input("PRESS ENTER TO CLOSE TERMINAL")

            try:
                row[2]
            except IndexError:
                print("ERROR: Font Bold does not exist in input.csv")
                print(
                    "ERROR: input.csv should be in form 'fontSize,fontName,fontBold(y or n),fontItalic (y or n)")
                input("PRESS ENTER TO CLOSE TERMINAL")

            try:
                row[3]
            except IndexError:
                print("ERROR: Font Italic does not exist in input.csv")
                print(
                    "ERROR: input.csv should be in form 'fontSize,fontName,fontBold(y or n),fontItalic (y or n)")
                input("PRESS ENTER TO CLOSE TERMINAL")

            textValuesDict['fontSize'] = row[0]
            textValuesDict['fontName'] = row[1]

            if (row[2] == 'y'):
                textValuesDict['fontBold'] = True
            elif (row[2] == 'n'):
                textValuesDict['fontBold'] = False
            else:
                print(
                    "WARNING: Bold value in CSV file is invalid. using default value of False")
                textValuesDict['fontBold'] = False

            if (row[3] == 'y'):
                textValuesDict['fontItalic'] = True
            elif (row[3] == 'n'):
                textValuesDict['fontItalic'] = False
            else:
                print(
                    "WARNING: Italic value in CSV file is invalid. using default value of False")
                textValuesDict['fontItalic'] = False

    return textValuesDict
