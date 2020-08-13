#include <OpenXLSX.hpp>
#include <iostream>

using namespace std;
using namespace OpenXLSX;

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #01: Basic Usage\n";
    cout << "********************************************************************************\n";

    XLDocument doc;
    doc.create("./Demo01.xlsx");
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell(XLCellReference("A1")).value() = 3.14159;
    wks.cell(XLCellReference("B1")).value() = 42;
    wks.cell(XLCellReference("C1")).value() = "  Hello OpenXLSX!  ";
    wks.cell(XLCellReference("D1")).value() = true;
    wks.cell(XLCellReference("E1")).value() = wks.cell(XLCellReference("C1")).value();

    auto A1 = wks.cell(XLCellReference("A1")).value();
    auto B1 = wks.cell(XLCellReference("B1")).value();
    auto C1 = wks.cell(XLCellReference("C1")).value();
    auto D1 = wks.cell(XLCellReference("D1")).value();
    auto E1 = wks.cell(XLCellReference("E1")).value();

    auto valueTypeAsString = [](OpenXLSX::XLValueType type) {
        switch (type) {
            case XLValueType::String:
                return "string";

            case XLValueType::Boolean:
                return "boolean";

            case XLValueType::Empty:
                return "empty";

            case XLValueType::Error:
                return "error";

            case XLValueType::Float:
                return "float";

            case XLValueType::Integer:
                return "integer";

            default:
                return "";
        }
    };

    cout << "Cell A1: (" << valueTypeAsString(A1.valueType()) << ") " << A1.get<double>() << endl;
    cout << "Cell B1: (" << valueTypeAsString(B1.valueType()) << ") " << B1.get<int>() << endl;
    cout << "Cell C1: (" << valueTypeAsString(C1.valueType()) << ") " << C1.get<std::string>() << endl;
    cout << "Cell D1: (" << valueTypeAsString(D1.valueType()) << ") " << D1.get<bool>() << endl;
    cout << "Cell E1: (" << valueTypeAsString(E1.valueType()) << ") " << E1.get<std::string_view>() << endl;

    doc.save();

    return 0;
}