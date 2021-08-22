import random

from OpenXLSX import XLDocument, XLCellReference

print('********************************************************************************')
print('DEMO PROGRAM #05: Ranges and Iterators')
print('********************************************************************************')

print()
print('Generating spreadsheet (1,048,576 rows x 8 columns) ...')

doc = XLDocument()
doc.create('Demo5.xlsx')
wks = doc.workbook().worksheet('Sheet1')
rng = wks.range(XLCellReference('A1'), XLCellReference(1048576, 8))

for cell in rng:
    cell.integerValue = random.randint(0, 99)

print('Saving spreadsheet (1,048,576 rows x 8 columns) ...')
doc.save()
doc.close()

print('Re-opening spreadsheet (1,048,576 rows x 8 columns) ...')
doc.open('Demo5.xlsx')
wks = doc.workbook().worksheet('Sheet1')
rng = wks.range(XLCellReference('A1'), XLCellReference(1048576, 8))

sum = 0
count = 0
print('Reading data from spreadsheet (1,048,576 rows x 8 columns) ...')
for cell in rng:
    count += 1
print('Cell count:', count)
for cell in rng:
    sum += cell.integerValue
print('Sum of cell values:', sum)

doc.close()
