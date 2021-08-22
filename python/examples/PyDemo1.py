from OpenXLSX import XLDocument, XLCellReference

print('********************************************************************************')
print('DEMO PROGRAM #01: Basic Usage')
print('********************************************************************************')

doc = XLDocument()
doc.create('Demo1.xlsx')
wks = doc.workbook().worksheet('Sheet1')

wks.cell('A1').floatValue = 3.14159
wks.cell(XLCellReference('B1')).integerValue = 42
wks.cell(XLCellReference('C1')).stringValue = '   Hello OpenXLSX!   '
wks.cell(XLCellReference('D1')).booleanValue = True
wks.cell(XLCellReference('E1')).value = wks.cell(XLCellReference('C1')).value

A1 = wks.cell(XLCellReference('A1')).value
B1 = wks.cell(XLCellReference('B1')).value
C1 = wks.cell(XLCellReference('C1')).value
D1 = wks.cell(XLCellReference('D1')).value
E1 = wks.cell(XLCellReference('E1')).value

print('Cell A1: (' + str(A1.valueType()) + ')', A1.floatValue)
print('Cell B1: (' + str(B1.valueType()) + ')', B1.integerValue)
print('Cell C1: (' + str(C1.valueType()) + ')', C1.stringValue)
print('Cell D1: (' + str(D1.valueType()) + ')', D1.booleanValue)
print('Cell E1: (' + str(E1.valueType()) + ')', E1.stringValue)

doc.save()
