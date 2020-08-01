import numpy as np
import os
import re
from win32com.client import Dispatch
from scipy import stats
from statsmodels.stats.proportion import proportions_ztest

def selection_sort(x):
    for i in range(len(x)):
        swap = i + np.argmin(x[i:])
        (x[i], x[swap]) = (x[swap], x[i])
    return x

class CIBC_PUC_V2(object):
    def __init__(self, path):
        self.xl = Dispatch("Excel.Application")
        self.xl.Visible = True
        self.wb = self.xl.Workbooks.Open(path)

    def setActivSheet(self, sheetName):
        return self.wb.Worksheets[sheetName]



    def findMyBase(self,baseName): #return row and column
        #Kak stava da znae s koj sheet raboti NZ !!!!! Towa ve magiq
        for column in range(1,10):
            #First check the "standart row for the base - 14 "
            var = sheet.Cells(14, column).Value
            if isinstance(var, str):
                if var.lower().strip() == baseName.lower().strip():
                    return 14, column
                    break
            else:
                FileFlag = False
                while True:
                    for row in range(1, 100):
                        if sheet.Cells(row, column).Value != None:
                            var = sheet.Cells(row, column).Value
                            if isinstance(var, str):
                                if var.lower().strip() == baseName.lower().strip():
                                    return row, column
                                    FileFlag = True
                                    break
                            else:
                                continue
                        else:
                            continue

                    if FileFlag or row >= 99:
                        # return set a error/warning message
                        break

    def compareBigerNumber(self,array):
        firstNumber = array[0]
        secondNumber = array[1]

        if firstNumber > secondNumber:
            return 0
        else:
            return 1


    def getZscore(self, baseRow1, baseRow2, p1, p2):
        count = np.array([p1, p2])
        nobs = np.array([baseRow1, baseRow2])
        stat, pval =  proportions_ztest(count, nobs)
        stat = abs(stat)
        return stat

    def scoreCheck(self, score, levels = [95, 90]):
        ret = 0 #NoSig Test of any Level
        #set the confidence level to a spesific sig level sig : conf
        confidenceLevels = {
            68 : 1.00,
            80 : 1.28,
            90 : 1.64,
            95 : 1.96,
            99 : 2.54
        }
        temp = 1
        for lv in selection_sort(levels):
            print(score)
            if float(score) > confidenceLevels[int(lv)]:
                ret = temp #Sig Test for the low sig level
                temp = temp + 1
            else:
                break
        return ret # Връща за кое ниво имаме тестване (0 - никое, 1 - малка буква (по-малкия процент) , 2-голяма (по-големия процент))


    def calculate_write_SigTest(self, baseRow1, baseRow2, p1, p2, Columnletters):
        lv = self.scoreCheck(self.getZscore(baseRow1, baseRow2, p1, p2))
        print(lv)
        


    #Still on progress (must be added the labels of the)
    def amIproperStatment(self, row, column):
        if sheet.Cells(row, column).Value != None:
            return True
        else:
            return False

    def findHowManyRows(self, column):
        emptyrows = 0
        myrow = 1
        while True:
            myrow = myrow + 1
            emptyrows = emptyrows + 1

            if self.amIproperStatment(myrow, column):
                emptyrows = 0

            #amIproperStatment(row = myrow, column = 1)
            if myrow > 100 or emptyrows >= 5:
                break

        return myrow - emptyrows


    def returnPercentForCategory(self,baseRow, baseColumn, initialColumn, vsColumn, endRow):
        def getStep(arrayOfRows, initialColumn): #Find the step/difference bettwen the row of the category and the row of the percent
            checkRow = arrayOfRows[0]

            if sheet.Cells(checkRow, initialColumn).Value > 0.0 and sheet.Cells(checkRow, initialColumn).Value < 1.0: #Most likly it's percent
                return 0
            else:
                dif = arrayOfRows[1] - arrayOfRows [0]
                if dif == 1:
                    return 0
                elif dif == 2:
                    return 1
                elif dif == 3:
                    return 1
                else:
                    #Print Warning
                    print("Warnign ! : getStep | difference must be bettwen 1-3")

        tempRow = baseRow + 1
        arrRow = np.array([])
        for row in range(tempRow, endRow):
            if sheet.Cells(row, baseColumn).Value != None:
                #print(sheet.Cells(row, baseColumn))
                arrRow = np.append(arrRow, [row])

        step = getStep(arrRow, initialColumn)
        #print(" ".join(str(x) for x in arrRow))
        #print(step)

        return arrRow, step


    def findSigTestColumn(self, Letter, baseRow, baseColumn):
        row = baseRow - 1
        column = baseColumn + 1

        if re.search("([A-Z])",  str(sheet.Cells(row, column))) != None:   #return None if fail (there is no letter in thes cell)
            while True:
                if re.search("(["+Letter.upper()+"])", str(sheet.Cells(row, column))) != None:
                    return column
                    break
                else:
                    column = column + 1

                if column > 100:
                    print("Warning please check -> Function: findSigTestColumn -> iterate over 100 column, no proper letter  ("+Letter+") found -> Row " + str(row))
                    break
        else:
            row = 1
            column = 2
            if re.search("([A-Z])",  str(sheet.Cells(row, column))) != None:
                pass # Да се допише обхождане за да намери реда на буквите за тестване
            

    def cptSigTest(self, initialColumn, vsColumn, BaseName): # 'vsAll - run the initial column vs all columns(with Letters in)'

        baseRow, baseColumn = myExcel.findMyBase(BaseName) # da go mahna ot tazi finkciq, za da ne lokiram bazata pri vqsko izpylnenie 
        letters = [initialColumn, vsColumn]
        print(letters)


        initialColumn = self.findSigTestColumn(initialColumn, baseRow, baseColumn)
        vsColumn = self.findSigTestColumn(vsColumn, baseRow, baseColumn)

        arrRows, step = self.returnPercentForCategory(baseRow, baseColumn, initialColumn, vsColumn, self.findHowManyRows(baseColumn))

        for row in arrRows + step:

            

            base1 = sheet.Cells(baseRow, initialColumn).Value
            base2 = sheet.Cells(baseRow, vsColumn).Value

            p1 = sheet.Cells(row, initialColumn).Value * base1
            p2 = sheet.Cells(row, vsColumn).Value * base2

            #print(str(p1) + "_"+str(p2)+"_"+str(base1)+"_"+str(base2))

            self.calculate_write_SigTest(base1, base2, p1,p2, letters)    

        print(initialColumn)
        print(vsColumn)


    def Close(self):
        #self.wb.Save
        self.wb.Close()
        self.xl.Quit()









# Calling the Function / Testing
myExcel = CIBC_PUC_V2(r'C:\Users\Danie\Desktop\Puc\PUC.xlsx')
sheet = myExcel.setActivSheet("PUC")
print(sheet.Cells(16,2))

#test = myExcel.scoreCheck( myExcel.getZscore(451,  475, 189, 290), [95, 90])
#print(test)

row, column =  myExcel.findMyBase("Base")
print(row, column)
myExcel.cptSigTest("A", "C", "Base")
#myExcel.returnPercentForCategory(14,1,7,8,50)

myExcel.Close()
