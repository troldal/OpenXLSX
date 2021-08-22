from OpenXLSX import XLDocument, XLCellReference

print('********************************************************************************')
print('DEMO PROGRAM #03: Unicode')
print('********************************************************************************')

doc1 = XLDocument()
doc1.create('Demo3.xlsx')
wks1 = doc1.workbook().worksheet('Sheet1')

wks1.cell(XLCellReference('A1')).stringValue = "안녕하세요 세계!"
wks1.cell(XLCellReference('A2')).stringValue = "你好，世界!"
wks1.cell(XLCellReference('A3')).stringValue = "こんにちは世界"
wks1.cell(XLCellReference('A4')).stringValue = "नमस्ते दुनिया!"
wks1.cell(XLCellReference('A5')).stringValue = "Привет, мир!"
wks1.cell(XLCellReference('A6')).stringValue = "Γειά σου Κόσμε!"

doc1.save()
doc1.saveAs('スプレッドシート.xlsx')
doc1.close()

doc2 = XLDocument()
doc2.open('Demo3.xlsx')
wks2 = doc2.workbook().worksheet('Sheet1')

print('Cell A1 (Korean)  :', wks2.cell(XLCellReference('A1')).stringValue)
print('Cell A2 (Chinese) :', wks2.cell(XLCellReference('A2')).stringValue)
print('Cell A3 (Japanese):', wks2.cell(XLCellReference('A3')).stringValue)
print('Cell A4 (Hindi)   :', wks2.cell(XLCellReference('A4')).stringValue)
print('Cell A5 (Russian) :', wks2.cell(XLCellReference('A5')).stringValue)
print('Cell A6 (Greek)   :', wks2.cell(XLCellReference('A6')).stringValue)

print()
print('NOTE: If you are using a Windows terminal, the above output will look like gibberish,\n'
      'because the Windows terminal does not support UTF-8 at the moment. To view to output,\n'
      'open the Demo03.xlsx file in Excel.')

doc2.close()
