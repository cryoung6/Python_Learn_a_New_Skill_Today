''' writeSumTotalsToFile.py

    Update the Excel worksheet object "ws" with the sum totals
    On lines 1 through 7 using the helper function writeData()

'''


def writeData(ws, name, row, theSum):
    ''' the hyperlink will be:
        Business_Expenses.xlsx#'Assets'!A2
        '''
    theCell = '!A2'
    file = "Business_Expenses.xlsx#'"
    ws.cell(row, 3).value = name
    ws.cell(row, 2).value = float(round(theSum, 2))
    theLink = file + name + "'" + theCell
    ws.cell(row, 2).hyperlink = theLink


def writeSumTotalsToFile(ws, sumAdv, sumProf, sumOther, sumOffice,
                         sumAssets, sumSupplies, sumMisc):

    writeData(ws, 'Advertising', 1, sumAdv)
    writeData(ws, 'Professional', 2, sumProf)
    writeData(ws, 'Other', 3, sumOther)
    writeData(ws, 'Office Expense', 4, sumOffice)
    writeData(ws, 'Assets', 5, sumAssets)
    writeData(ws, 'Supplies', 6, sumSupplies)
    writeData(ws, 'Miscellaneous', 7, sumMisc)

    return ws
