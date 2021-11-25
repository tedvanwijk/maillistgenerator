import numpy as np
import pandas as pd

def findName(name, shifts):
    for i in range(len(shifts)):
        try:
            if shifts[i]['name'] == name:
                return i
        except:
            return None
    return None

def addShift(name,shift,shifts):
    if name == '':
        return shifts
    nameIndex = findName(name,shifts)
    # if the name does not yet exits, we make a new entry with the new shift in it
    if nameIndex == None:
        newPerson = {
            "name": name,
            "shifts": [
                shift
            ]
        }
        shifts = np.append(shifts,newPerson)
    # if it does exist, we simply add the new shift to the shifts entry of this person
    else:
        currentPerson = shifts[nameIndex]
        for i in range(len(currentPerson['shifts'])):
            currentShift = currentPerson['shifts'][i]
            if currentShift['day'] == shift['day'] and currentShift['startTime'] < shift['endTime'] and currentShift['endTime'] > shift['startTime']:
                minTime = np.min([currentShift['startTime'], shift['startTime']])
                maxTime = np.max([currentShift['endTime'], shift['endTime']])
                print(f"\nLet op! Er is een dubbele shift ingepland voor {currentPerson['name']}: {currentShift['shift']} end {shift['shift']} op {currentShift['day']} van {minTime} tot {maxTime}\n")
        shifts[nameIndex]['shifts'] = np.append(shifts[nameIndex]['shifts'], shift)
    return shifts

def main():
    print('Hallo daar! Welkom bij de maillijst generator! Om dit te gebruiken moet je Python hebben geinstalleerd met de packages numpy en pandas. \n')

    fileName = input('Wat is de bestandsnaam? (incl. de bestandsextensie zoals .xlsx of .csv!): \n')
    print('\nJe wordt nu gevraagd om de namen van de sheets op te geven waar een maillijst van moet worden gemaakt. Dus, stel je hebt een sheet "Zaterdag" en "Zondag" dan vul je die een voor een in \n')

    sheetInput = '_'
    sheets = np.array([])

    while sheetInput != '':
        sheetInput = input('Vul de naam van een enkele sheet in. Als je ze allemaal in hebt gevuld, druk je alleen op enter: \n')
        sheets = np.append(sheets, sheetInput)
    sheets = sheets[:-1]

    outputFileName = input("Wat moet de naam worden van het uiteindelijke bestand? (zonder bestandsextensie, het wordt een .csv): \n")

    shifts = np.array([])

    for sheet in sheets:
        data = pd.read_excel(fileName, sheet_name=sheet)
        columns = data.columns
        startingColumn = columns.get_loc('-') + 1
        # the total length minus where the times start is the length of the time columns
        columnLoopLength = len(columns) - startingColumn
        for index, row in data.iterrows():
            previousName = ''
            newShift = {}
            for i in range(columnLoopLength):
                currentTime = columns[startingColumn + i]

                # if the current name is nan, it is the ending time of the previous person's shift
                if pd.isnull(row[currentTime]):
                    if previousName != '':
                        newShift['endTime'] = currentTime
                        shifts = addShift(previousName,newShift,shifts)
                        previousName = ''
                        continue
                    else:
                        continue
                endTime = columns[startingColumn + i + 1]
                currentName = row[currentTime].strip().lower()

                if previousName != currentName:
                    # if this is true that means we have reached the end of someone's shift, so we can add it to the shifts array
                    shifts = addShift(previousName,newShift,shifts)
                    # make new shift for the new person
                    newShift = {
                        "day": sheet,
                        "shift": row['Shift'],
                        "startTime": currentTime,
                        "endTime": endTime
                    }
                else:
                    newShift['endTime'] = endTime
                previousName = currentName

    export = []
    for i in range(len(shifts)):
        currentPerson = shifts[i]
        # sort shifts based on their start time but only for shifts that are on the same day
        for sheet in sheets:
            currentPerson['shifts'] = sorted(currentPerson['shifts'], key=lambda e: (e['day'] == sheet, e['startTime']))
        row = np.array(currentPerson['name'])
        shiftText = ''
        currentDay = ''
        for ii in range(len(currentPerson['shifts'])):
            currentShift = currentPerson['shifts'][ii]
            startTime = currentShift['startTime'].strftime("%H:%M")
            endTime = currentShift['endTime'].strftime("%H:%M")
            if currentShift['day'] != currentDay:
                shiftText += f"{currentShift['day']}: \n" 
                currentDay = currentShift['day']
            shiftText += f"{currentShift['shift']} van {startTime} t/m {endTime} \n"
        row = np.append(row, shiftText)
        export.append(list(row))

    export = np.asarray(export)
    df = pd.DataFrame(export)
    df.to_csv(f"{outputFileName}.csv")

    input('\nLet op! hierboven kunnen nog belangrijke meldingen staan! Druk op enter om het script af te sluiten...')

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print('\nEr is een foutmelding opgetreden. Als je deze error niet snapt of het programma werkt niet zoals het hoort, kan je een mailtje sturen naar tedvanwijk@gmail.com met onderstaande error erbij:\n')
        print(e)
        input()