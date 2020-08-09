#include <OpenXLSX.hpp>
#include <iostream>

using namespace std;
using namespace OpenXLSX;

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #04: Number Formats\n";
    cout << "********************************************************************************\n";

    XLDocument doc1;
    doc1.create("./Demo04.xlsx");
    auto wks1 = doc1.workbook().worksheet("Sheet1");

    wks1.cell("A1").value() = 0.01;
    wks1.cell("B1").value() = 0.02;
    wks1.cell("C1").value() = 0.03;
    wks1.cell("A2").value() = 0.001;
    wks1.cell("B2").value() = 0.002;
    wks1.cell("C2").value() = 0.003;
    wks1.cell("A3").value() = 1e-4;
    wks1.cell("B3").value() = 2e-4;
    wks1.cell("C3").value() = 3e-4;

    wks1.cell("A4").value() = 1;
    wks1.cell("B4").value() = 2;
    wks1.cell("C4").value() = 3;
    wks1.cell("A5").value() = 845287496;
    wks1.cell("B5").value() = 175397487;
    wks1.cell("C5").value() = 973853975;
    wks1.cell("A6").value() = 2e10;
    wks1.cell("B6").value() = 3e11;
    wks1.cell("C6").value() = 4e12;

    doc1.save();
    doc1.close();

    XLDocument doc2;
    doc2.open("./Demo04.xlsx");
    auto wks2 = doc2.workbook().worksheet("Sheet1");

    auto PrintCell = [](const XLCell& cell) {
        cout << "Cell type is ";

        switch (cell.valueType()) {
            case XLValueType::Empty:
                cout << "XLValueType::Empty";
                break;

            case XLValueType::Float:
                cout << "XLValueType::Float and the value is " << cell.value().get<double>() << endl;
                break;

            case XLValueType::Integer:
                cout << "XLValueType::Integer and the value is " << cell.value().get<int64_t>() << endl;
                break;

            default:
                cout << "Unknown";
        }
    };

    cout << "Cell A1: ";
    PrintCell(wks2.cell("A1"));

    cout << "Cell B1: ";
    PrintCell(wks2.cell("B1"));

    cout << "Cell C1: ";
    PrintCell(wks2.cell("C1"));

    cout << "Cell A2: ";
    PrintCell(wks2.cell("A2"));

    cout << "Cell B2: ";
    PrintCell(wks2.cell("B2"));

    cout << "Cell C2: ";
    PrintCell(wks2.cell("C2"));

    cout << "Cell A3: ";
    PrintCell(wks2.cell("A3"));

    cout << "Cell B3: ";
    PrintCell(wks2.cell("B3"));

    cout << "Cell C3: ";
    PrintCell(wks2.cell("C3"));

    cout << "Cell A4: ";
    PrintCell(wks2.cell("A4"));

    cout << "Cell B4: ";
    PrintCell(wks2.cell("B4"));

    cout << "Cell C4: ";
    PrintCell(wks2.cell("C4"));

    cout << "Cell A5: ";
    PrintCell(wks2.cell("A5"));

    cout << "Cell B5: ";
    PrintCell(wks2.cell("B5"));

    cout << "Cell C5: ";
    PrintCell(wks2.cell("C5"));

    cout << "Cell A6: ";
    PrintCell(wks2.cell("A6"));

    cout << "Cell B6: ";
    PrintCell(wks2.cell("B6"));

    cout << "Cell C6: ";
    PrintCell(wks2.cell("C6"));

    doc2.close();

    return 0;
}