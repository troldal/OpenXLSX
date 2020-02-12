#include <iostream>
#include <OpenXLSX/OpenXLSX.h>

using namespace std;
using namespace OpenXLSX;

int main() {

    XLDocument doc1;
    doc1.CreateDocument("./UnicodeTest.xlsx");
    auto wks1 = doc1.Workbook().Worksheet("Sheet1");

    wks1.Cell("A1").Value() = "안녕하세요 세계!";
    wks1.Cell("A2").Value() = "你好，世界!";
    wks1.Cell("A3").Value() = "こんにちは世界";
    wks1.Cell("A4").Value() = "नमस्ते दुनिया!";
    wks1.Cell("A5").Value() = "Привет, мир!";
    wks1.Cell("A6").Value() = "Γειά σου Κόσμε!";

    doc1.SaveDocument();
    doc1.CloseDocument();

    XLDocument doc2;
    doc2.OpenDocument("./UnicodeTest.xlsx");
    auto wks2 = doc2.Workbook().Worksheet("Sheet1");

    cout << "Cell A1: " << wks2.Cell("A1").Value().Get<std::string>() << endl;
    cout << "Cell A2: " << wks2.Cell("A2").Value().Get<std::string>() << endl;
    cout << "Cell A3: " << wks2.Cell("A3").Value().Get<std::string>() << endl;
    cout << "Cell A4: " << wks2.Cell("A4").Value().Get<std::string>() << endl;
    cout << "Cell A5: " << wks2.Cell("A5").Value().Get<std::string>() << endl;
    cout << "Cell A6: " << wks2.Cell("A6").Value().Get<std::string>() << endl;

    doc2.CloseDocument();

    return 0;
}