import numpy as np
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
                        # return set a erro/warning message
                        break



    def getZscore(self, baseRow1, baseRow2, p1, p2):
        count = np.array([p1, p2])
        nobs = np.array([baseRow1, baseRow2])
        stat, pval =  proportions_ztest(count, nobs)
        stat = abs(stat)
        return stat

    def scoreCheck(self, score, levels):
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
        return ret

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

    def Close(self):
        #self.wb.Save
        self.wb.Close()
        #self.xl.Quit()









# Calling the Function / Testing
myExcel = CIBC_PUC_V2(r"C:\Users\Daniel.Mladenov\Desktop\PUK\PUC.xlsx")
sheet = myExcel.setActivSheet("0001")
print(sheet.Cells(13,2))

#test = myExcel.scoreCheck( myExcel.getZscore(200, 220, 100, 140), [95, 90])
#print(test)

row, column =  myExcel.findMyBase("Base")
print(row, column)
myExcel.Close()
