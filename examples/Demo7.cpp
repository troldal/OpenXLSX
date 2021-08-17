#include <OpenXLSX.hpp>
#include <iostream>
#include <random>
#include <deque>
#include <list>

using namespace std;
using namespace OpenXLSX;

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #07: Row Handling - Implicit Conversion\n";
    cout << "********************************************************************************\n";

    cout << "\nGenerating spreadsheet ..." << endl;
    XLDocument doc;
    doc.create("./Demo07.xlsx");
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.row(1).values() = std::vector<int>{1,2,3,4,5,6,7,8};
    wks.row(2).values() = std::vector<double>{1.1, 2.2, 3.3, 4.4, 5.5, 6.6, 7.7, 8.8};
    wks.row(3).values() = std::vector<bool>{true, false, true, false, true, false, true, false};
    wks.row(4).values() = std::vector<std::string>{"A", "B", "C", "D", "E", "F", "G", "H"};

    cout << "Saving spreadsheet ..." << endl;
    doc.save();
    doc.close();

    cout << "Re-opening spreadsheet ..." << endl << endl;
    doc.open("./Demo07.xlsx");
    wks = doc.workbook().worksheet("Sheet1");

    cout << "Conversion to std::vector<XLCellValue> ..." << endl;
    for (auto& row : wks.rows()) {
        for (auto& value : std::vector<XLCellValue>(row.values())) {
            cout << value << " ";
        }
        cout << endl;
    }

    cout << endl << "Conversion to std::deque<XLCellValue> ..." << endl;
    for (auto& row : wks.rows()) {
        for (auto& value : std::deque<XLCellValue>(row.values())) {
            cout << value << " ";
        }
        cout << endl;
    }

    cout << endl << "Conversion to std::list<XLCellValue> ..." << endl;
    for (auto& row : wks.rows()) {
        for (auto& value : std::list<XLCellValue>(row.values())) {
            cout << value << " ";
        }
        cout << endl;
    }

    cout << endl << "Conversion to std::vector<[int, double, bool, std::string]> ..." << endl;
    for (auto& value : std::vector<int>(wks.row(1).values())) cout << value << " ";
    cout << endl;
    for (auto& value : std::vector<double>(wks.row(2).values())) cout << value << " ";
    cout << endl;
    for (bool value : std::vector<bool>(wks.row(3).values())) cout << value << " ";
    cout << endl;
    for (auto& value : std::vector<std::string>(wks.row(4).values())) cout << value << " ";
    cout << endl;

    cout << endl << "Conversion to std::deque<[int, double, bool, std::string]> ..." << endl;
    for (auto& value : std::deque<int>(wks.row(1).values())) cout << value << " ";
    cout << endl;
    for (auto& value : std::deque<double>(wks.row(2).values())) cout << value << " ";
    cout << endl;
    for (auto& value : std::deque<bool>(wks.row(3).values())) cout << value << " ";
    cout << endl;
    for (auto& value : std::deque<std::string>(wks.row(4).values())) cout << value << " ";
    cout << endl;

    cout << endl << "Conversion to std::list<[int, double, bool, std::string]> ..." << endl;
    for (auto& value : std::list<int>(wks.row(1).values())) cout << value << " ";
    cout << endl;
    for (auto& value : std::list<double>(wks.row(2).values())) cout << value << " ";
    cout << endl;
    for (auto& value : std::list<bool>(wks.row(3).values())) cout << value << " ";
    cout << endl;
    for (auto& value : std::list<std::string>(wks.row(4).values())) cout << value << " ";
    cout << endl;

    doc.close();

    return 0;
}
