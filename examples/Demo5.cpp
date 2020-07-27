#include <iostream>
#include <OpenXLSX.hpp>

using namespace std;
using namespace OpenXLSX;

int main()
{
    XLDocument doc1;
    doc1.create("./NumberTest.xlsx");
    auto wks1 = doc1.workbook().worksheet("Sheet1");

    wks1.cell("A1").value() = 0.01;
    wks1.cell("B1").value() = 0.02;
    wks1.cell("C1").value() = 0.03;
    wks1.cell("A2").value() = 0.001;
    wks1.cell("B2").value() = 0.002;
    wks1.cell("C2").value() = 0.003;

    doc1.save();
    doc1.close();

    XLDocument doc2;
    doc2.open("./NumberTest.xlsx");
    auto wks2 = doc2.workbook().worksheet("Sheet1");

    cout << "Cell A1: " << wks2.cell("A1").value().get<double>() << endl;
    cout << "Cell B1: " << wks2.cell("B1").value().get<double>() << endl;
    cout << "Cell C1: " << wks2.cell("C1").value().get<double>() << endl;
    cout << "Cell A2: " << wks2.cell("A2").value().get<double>() << endl;
    cout << "Cell B2: " << wks2.cell("B2").value().get<double>() << endl;
    cout << "Cell C2: " << wks2.cell("C2").value().get<double>() << endl;

    auto PrintCell = [](const XLCell& cell) {
        cout << "Cell type is ";

        switch (cell.valueType()) {
            case XLValueType::Empty:
                cout << "XLValueType::Empty";
                break;

            case XLValueType::Float:
                cout << "XLValueType::Float and the value is " << cell.value().get<float>() << endl;
                break;

            case XLValueType::Integer:
                cout << "XLValueType::Integer and the value is " << cell.value().get<int>() << endl;
                break;

            default:
                cout << "Unknown";
        }
    };

    PrintCell(wks2.cell("A1"));
    PrintCell(wks2.cell("B1"));
    PrintCell(wks2.cell("C1"));
    PrintCell(wks2.cell("A2"));
    PrintCell(wks2.cell("B2"));
    PrintCell(wks2.cell("C2"));

    doc2.close();

    return 0;
}