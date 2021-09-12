from OpenXLSX import XLDocument, XLValueType

print('********************************************************************************')
print('DEMO PROGRAM #04: Number Formats')
print('********************************************************************************')

doc1 = XLDocument()
doc1.create('Demo4.xlsx')
wks1 = doc1.workbook().worksheet('Sheet1')

wks1.cell('A1').floatValue = 0.01
wks1.cell('B1').floatValue = 0.02
wks1.cell('C1').floatValue = 0.03
wks1.cell('A2').floatValue = 0.001
wks1.cell('B2').floatValue = 0.002
wks1.cell('C2').floatValue = 0.003
wks1.cell('A3').floatValue = 1e-4
wks1.cell('B3').floatValue = 2e-4
wks1.cell('C3').floatValue = 3e-4

wks1.cell('A4').integerValue = 1
wks1.cell('B4').integerValue = 2
wks1.cell('C4').integerValue = 3
wks1.cell('A5').integerValue = 845287496
wks1.cell('B5').integerValue = 175397487
wks1.cell('C5').integerValue = 973853975
wks1.cell('A6').integerValue = int(2e10)
wks1.cell('B6').integerValue = int(3e11)
wks1.cell('C6').integerValue = int(4e12)

doc1.save()
doc1.close()

doc2 = XLDocument()
doc2.open('Demo4.xlsx')
wks2 = doc2.workbook().worksheet('Sheet1')


def printCell(cell):
    if cell.valueType() == XLValueType.Empty:
        print('Cell type is XLValueType.Empty')
        return

    if cell.valueType() == XLValueType.Float:
        print('Cell type is XLValueType.Float and the getValue is', cell.floatValue)
        return

    if cell.valueType() == XLValueType.Integer:
        print('Cell type is XLValueType.Integer and the getValue is', cell.integerValue)
        return


print('Cell A1: ', end='')
printCell(wks2.cell('A1'))

print('Cell B1: ', end='')
printCell(wks2.cell('B1'))

print('Cell C1: ', end='')
printCell(wks2.cell('C1'))

print('Cell A2: ', end='')
printCell(wks2.cell('A2'))

print('Cell B2: ', end='')
printCell(wks2.cell('B2'))

print('Cell C2: ', end='')
printCell(wks2.cell('C2'))

print('Cell A3: ', end='')
printCell(wks2.cell('A3'))

print('Cell B3: ', end='')
printCell(wks2.cell('B3'))

print('Cell C3: ', end='')
printCell(wks2.cell('C3'))

print('Cell A4: ', end='')
printCell(wks2.cell('A4'))

print('Cell B4: ', end='')
printCell(wks2.cell('B4'))

print('Cell C4: ', end='')
printCell(wks2.cell('C4'))

print('Cell A5: ', end='')
printCell(wks2.cell('A5'))

print('Cell B5: ', end='')
printCell(wks2.cell('B5'))

print('Cell C5: ', end='')
printCell(wks2.cell('C5'))

print('Cell A6: ', end='')
printCell(wks2.cell('A6'))

print('Cell B6: ', end='')
printCell(wks2.cell('B6'))

print('Cell C6: ', end='')
printCell(wks2.cell('C6'))
