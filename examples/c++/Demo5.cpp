#include <iostream>
#include <OpenXLSX/OpenXLSX.h>

using namespace std;
using namespace OpenXLSX;

int main() {

    XLDocument doc1;
    doc1.CreateDocument("./NumberTest.xlsx");
    auto wks1 = doc1.Workbook().Worksheet("Sheet1");

    wks1.Cell("A1").Value() = 0.01;
    wks1.Cell("B1").Value() = 0.02;
    wks1.Cell("C1").Value() = 0.03;
    wks1.Cell("A2").Value() = 0.001;
    wks1.Cell("B2").Value() = 0.002;
    wks1.Cell("C2").Value() = 0.003;

    doc1.SaveDocument();
    doc1.CloseDocument();

    XLDocument doc2;
    doc2.OpenDocument("./NumberTest.xlsx");
    auto wks2 = doc2.Workbook().Worksheet("Sheet1");

    cout << "Cell A1: " << wks2.Cell("A1").Value().Get<double>() << endl;
    cout << "Cell B1: " << wks2.Cell("B1").Value().Get<double>() << endl;
    cout << "Cell C1: " << wks2.Cell("C1").Value().Get<double>() << endl;
    cout << "Cell A2: " << wks2.Cell("A2").Value().Get<double>() << endl;
    cout << "Cell B2: " << wks2.Cell("B2").Value().Get<double>() << endl;
    cout << "Cell C2: " << wks2.Cell("C2").Value().Get<double>() << endl;

    auto PrintCell = [](const XLCell& cell) {

        cout << "Cell type is ";

        switch (cell.ValueType()) {
            case XLValueType::Empty:
                cout << "XLValueType::Empty";
                break;

            case XLValueType::Float:
                cout << "XLValueType::Float and the value is " << cell.Value().Get<float>() << endl;
                break;

            case XLValueType::Integer:
                cout << "XLValueType::Integer and the value is " << cell.Value().Get<int>() << endl;
                break;

            default:
                cout << "Unknown";
        }

    };

    PrintCell(wks2.Cell("A1"));
    PrintCell(wks2.Cell("B1"));
    PrintCell(wks2.Cell("C1"));
    PrintCell(wks2.Cell("A2"));
    PrintCell(wks2.Cell("B2"));
    PrintCell(wks2.Cell("C2"));

    doc2.CloseDocument();

    return 0;
}