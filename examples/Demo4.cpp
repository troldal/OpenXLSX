#include <iostream>
#include <OpenXLSX.hpp>

using namespace std;
using namespace OpenXLSX;

int main()
{
    XLDocument doc1;
    doc1.create("./UnicodeTest.xlsx");
    auto wks1 = doc1.workbook()->Worksheet("Sheet1");

    wks1.Cell(XLCellReference("A1")).Value() = "안녕하세요 세계!";
    wks1.Cell(XLCellReference("A2")).Value() = "你好，世界!";
    wks1.Cell(XLCellReference("A3")).Value() = "こんにちは世界";
    wks1.Cell(XLCellReference("A4")).Value() = "नमस्ते दुनिया!";
    wks1.Cell(XLCellReference("A5")).Value() = "Привет, мир!";
    wks1.Cell(XLCellReference("A6")).Value() = "Γειά σου Κόσμε!";

    doc1.save();
    doc1.close();

    XLDocument doc2;
    doc2.open("./UnicodeTest.xlsx");
    auto wks2 = doc2.workbook()->Worksheet("Sheet1");

    cout << "Cell A1: " << wks2.Cell(XLCellReference("A1")).Value().Get<std::string>() << endl;
    cout << "Cell A2: " << wks2.Cell(XLCellReference("A2")).Value().Get<std::string>() << endl;
    cout << "Cell A3: " << wks2.Cell(XLCellReference("A3")).Value().Get<std::string>() << endl;
    cout << "Cell A4: " << wks2.Cell(XLCellReference("A4")).Value().Get<std::string>() << endl;
    cout << "Cell A5: " << wks2.Cell(XLCellReference("A5")).Value().Get<std::string>() << endl;
    cout << "Cell A6: " << wks2.Cell(XLCellReference("A6")).Value().Get<std::string>() << endl;

    doc2.close();

    return 0;
}