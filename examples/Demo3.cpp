#include <OpenXLSX.hpp>
#include <iostream>

using namespace std;
using namespace OpenXLSX;

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #03: Unicode\n";
    cout << "********************************************************************************\n";

    XLDocument doc1;
    doc1.create("./Demo03.xlsx");
    auto wks1 = doc1.workbook().worksheet("Sheet1");

    wks1.cell(XLCellReference("A1")).value() = "안녕하세요 세계!";
    wks1.cell(XLCellReference("A2")).value() = "你好，世界!";
    wks1.cell(XLCellReference("A3")).value() = "こんにちは世界";
    wks1.cell(XLCellReference("A4")).value() = "नमस्ते दुनिया!";
    wks1.cell(XLCellReference("A5")).value() = "Привет, мир!";
    wks1.cell(XLCellReference("A6")).value() = "Γειά σου Κόσμε!";

    doc1.save();
    doc1.saveAs("./スプレッドシート.xlsx");
    doc1.close();

    XLDocument doc2;
    doc2.open("./Demo03.xlsx");
    auto wks2 = doc2.workbook().worksheet("Sheet1");

    cout << "Cell A1 (Korean)  : " << wks2.cell(XLCellReference("A1")).value().get<std::string>() << endl;
    cout << "Cell A2 (Chinese) : " << wks2.cell(XLCellReference("A2")).value().get<std::string>() << endl;
    cout << "Cell A3 (Japanese): " << wks2.cell(XLCellReference("A3")).value().get<std::string>() << endl;
    cout << "Cell A4 (Hindi)   : " << wks2.cell(XLCellReference("A4")).value().get<std::string>() << endl;
    cout << "Cell A5 (Russian) : " << wks2.cell(XLCellReference("A5")).value().get<std::string>() << endl;
    cout << "Cell A6 (Greek)   : " << wks2.cell(XLCellReference("A6")).value().get<std::string>() << endl;


    cout << "\nNOTE: If you are using a Windows terminal, the above output will look like gibberish,\n"
            "because the Windows terminal does not support UTF-8 at the moment. To view to output,\n"
            "open the Demo03.xlsx file in Excel.\n\n";

    doc2.close();

    return 0;
}