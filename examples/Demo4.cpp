#include <iostream>
#include <OpenXLSX.hpp>

using namespace std;
using namespace OpenXLSX;

int main()
{
    XLDocument doc1;
    doc1.create("./UnicodeTest.xlsx");
    auto wks1 = doc1.workbook().worksheet("Sheet1");

    wks1.cell(XLCellReference("A1")).value() = "안녕하세요 세계!";
    wks1.cell(XLCellReference("A2")).value() = "你好，世界!";
    wks1.cell(XLCellReference("A3")).value() = "こんにちは世界";
    wks1.cell(XLCellReference("A4")).value() = "नमस्ते दुनिया!";
    wks1.cell(XLCellReference("A5")).value() = "Привет, мир!";
    wks1.cell(XLCellReference("A6")).value() = "Γειά σου Κόσμε!";

    doc1.save();
    doc1.close();

    XLDocument doc2;
    doc2.open("./UnicodeTest.xlsx");
    auto wks2 = doc2.workbook().worksheet("Sheet1");

    cout << "Cell A1: " << wks2.cell(XLCellReference("A1")).value().get<std::string>() << endl;
    cout << "Cell A2: " << wks2.cell(XLCellReference("A2")).value().get<std::string>() << endl;
    cout << "Cell A3: " << wks2.cell(XLCellReference("A3")).value().get<std::string>() << endl;
    cout << "Cell A4: " << wks2.cell(XLCellReference("A4")).value().get<std::string>() << endl;
    cout << "Cell A5: " << wks2.cell(XLCellReference("A5")).value().get<std::string>() << endl;
    cout << "Cell A6: " << wks2.cell(XLCellReference("A6")).value().get<std::string>() << endl;

    doc2.close();

    return 0;
}