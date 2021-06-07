# categoryWorksheets.py
'''
    The four lists have the same format
    advList, deprList, otherList, profList

    advList = [[thisDate, thisDesc, thisAmt, thisCategory],
               [thisDate, thisDesc, thisAmt, thisCategory]]

    1.  create 4 worksheets
    2.  write the list data to the respective worksheets
        Column 1 = thisDate
        Column 2 = thisDesc
        Column 3 = thisAmt

'''
from formatExcelFile import formatCategoryWs


def categoryWorksheets(wb, advList, otherList, profList, officeList,
                       assetsList, suppliesList, miscList):
    wsAdv = wb.create_sheet('Advertising')
    wsOther = wb.create_sheet('Other')
    wsProf = wb.create_sheet('Professional')
    wsOffice = wb.create_sheet('Office Supplies')
    wsAssets = wb.create_sheet('Assets')
    wsSupplies = wb.create_sheet('Supplies')
    wsMisc = wb.create_sheet('Miscellaneous')

    list_2_worksheet(advList, wsAdv)
    list_2_worksheet(otherList, wsOther)
    list_2_worksheet(profList, wsProf)
    list_2_worksheet(officeList, wsOffice)
    list_2_worksheet(assetsList, wsAssets)
    list_2_worksheet(suppliesList, wsSupplies)
    list_2_worksheet(miscList, wsMisc)

    wsAdv = formatCategoryWs(wsAdv, advList)
    wsOther = formatCategoryWs(wsOther, otherList)
    wsProf = formatCategoryWs(wsProf, profList)
    wsAssets = formatCategoryWs(wsAssets, assetsList)
    wsOffice = formatCategoryWs(wsOffice, officeList)
    wsSupplies = formatCategoryWs(wsSupplies, suppliesList)
    wsMisc = formatCategoryWs(wsMisc, miscList)
    return wb


def list_2_worksheet(theList, ws):
    items = len(theList)
    outRow = 2
    for item in range(items):
        for col in range(3):
            ws.cell(row=outRow, column=col + 1).value = theList[item][col]
        outRow += 1


# testing ================================================================
#from openpyxl import Workbook
#wb = Workbook()
#otherList = [['10/04/20', 'Book Proof', '11.40', 'OTHER PUBLISHING EXPENSES'],
#             ['06/14/20', 'paper', '7.65', 'OTHER PUBLISHING EXPENSES'],
#             ['04/14/20', 'Coreldraw', '23.90', 'OTHER PUBLISHING EXPENSES']]
#deprList = [['09/16/20', 'Apple Store keyboard', '31.44',
#             'DEPRECIATION & SECTION 179'],
#            ['11/11/20', 'Apple store pencil', '80.42',
#             'DEPRECIATION & SECTION 179']]
#advList = [['07/02/20', 'AMZ advertising', '76.41', 'ADVERTISING'],
#           ['02/02/20', 'AMZ Advertising', '97.18', 'ADVERTISING'],
#           ['11/02/20', 'AMZ Advertising', '133.65', 'ADVERTISING']]
#profList = [['01/08/20', 'Professional services CPA', '50.00',
#             'LEGAL & PROFESSIONAL']]
#
#wb = categoryWorksheets(wb, advList, deprList, otherList, profList)
#
#wb.save('testing.xlsx')
