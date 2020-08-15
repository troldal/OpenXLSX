import OpenXLSX

doc = OpenXLSX.XLDocument()
doc.create('Demo1.xlsx')
wks = doc.workbook().worksheet('Sheet1')

wks.cell(OpenXLSX.XLCellReference('A1')).value().setFloat(3.14159)
wks.cell(OpenXLSX.XLCellReference('B1')).value().setInteger(42)
wks.cell(OpenXLSX.XLCellReference('C1')).value().setString('   Hello OpenXLSX!   ')
wks.cell(OpenXLSX.XLCellReference('D1')).value().setBoolean(True)
wks.cell(OpenXLSX.XLCellReference('E1')).value().setString(wks.cell(OpenXLSX.XLCellReference('C1')).value().getString())

doc.save()
