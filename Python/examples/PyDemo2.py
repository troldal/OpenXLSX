from OpenXLSX import XLDocument, XLColor

print('********************************************************************************')
print('DEMO PROGRAM #02: Sheet Handling')
print('********************************************************************************')

doc = XLDocument()
doc.create('Demo2.xlsx')
wbk = doc.workbook()

print()
print('Sheets in workbook:')
for sheetName in wbk.worksheetNames():
    print(sheetName)

print()
print('Adding new sheet "MySheet01"')
wbk.addWorksheet('MySheet01')

print('Adding new sheet "MySheet02"')
wbk.addWorksheet('MySheet02')

print('cloning sheet "Sheet1" to new sheet "MySheet03"')
wbk.sheet('Sheet1').worksheet().clone('MySheet03')

print('cloning sheet "MySheet01" to new sheet "MySheet04"')
wbk.cloneSheet('MySheet01', 'MySheet04')

print()
print('Sheets in workbook:')
for sheetName in wbk.worksheetNames():
    print(sheetName)

print()
print('Deleting sheet "Sheet1"')
wbk.deleteSheet('Sheet1')

print('Moving sheet "MySheet04" to index 1')
wbk.worksheet('MySheet04').setIndex(1)

print('Moving sheet "MySheet03" to index 2')
wbk.worksheet('MySheet03').setIndex(2)

print('Moving sheet "MySheet02" to index 3')
wbk.worksheet('MySheet02').setIndex(3)

print('Moving sheet "MySheet01" to index 4')
wbk.worksheet('MySheet01').setIndex(4)

print()
print('Sheets in workbook:')
for sheetName in wbk.worksheetNames():
    print(sheetName)

wbk.sheet('MySheet01').setColor(XLColor(0, 0, 0))
wbk.sheet('MySheet02').setColor(XLColor(255, 0, 0))
wbk.sheet('MySheet03').setColor(XLColor(0, 255, 0))
wbk.sheet('MySheet04').setColor(XLColor(0, 0, 255))

doc.save()
