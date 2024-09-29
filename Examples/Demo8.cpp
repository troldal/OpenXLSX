#include <OpenXLSX.hpp>
#include <deque>
#include <iostream>
#include <list>
#include <random>

using namespace std;
using namespace OpenXLSX;

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #08: Row Data - Implicit Conversion\n";
    cout << "********************************************************************************\n";


    // The previous example showed, among other things, how a std::vector of XLCellValue objects
    // could be assigned to to an XLRow object, to populate the data to the cells in that row.
    // In fact, this can be done with any container supporting bi-directional iterators, holding
    // any data type that is convertible to a valid cell value. This is illustrated in this example.

    // First, create a new document and access the sheet named 'Sheet1'.
    cout << "\nGenerating spreadsheet ..." << endl;
    XLDocument doc;
    doc.create("./Demo08.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // A std::vector holding values that are convertible to a cell value can be assigned to an XLRow
    // object, using the 'values()' method. For example ints, doubles, bools and std::strings.
    wks.row(1).values() = std::vector<int> { 1, 2, 3, 4, 5, 6, 7, 8 };
    wks.row(2).values() = std::vector<double> { 1.1, 2.2, 3.3, 4.4, 5.5, 6.6, 7.7, 8.8 };
    wks.row(3).values() = std::vector<bool> { true, false, true, false, true, false, true, false };
    wks.row(4).values() = std::vector<std::string> { "A", "B", "C", "D", "E", "F", "G", "H" };

    // Save the sheet...
    cout << "Saving spreadsheet ..." << endl;
    doc.save();
    doc.close();

    // ...and reopen it (just to make sure that it is a valid .xlsx file)
    cout << "Re-opening spreadsheet ..." << endl << endl;
    doc.open("./Demo08.xlsx");
    wks = doc.workbook().worksheet("Sheet1");

    // The '.values()' method returns a proxy object that can be converted to any container supporting
    // bi-directional iterators. The following three blocks shows how the row data can be converted to
    // std::vector, std::deque, and std::list. In principle, this will also work with non-stl containers,
    // e.g. containers in the Qt framework. This has not been tested, though.

    // Conversion to std::vector<XLCellValue>
    cout << "Conversion to std::vector<XLCellValue> ..." << endl;
    for (auto& row : wks.rows()) {
        for (auto& value : std::vector<XLCellValue>(row.values())) {
            cout << value << " ";
        }
        cout << endl;
    }

    // Conversion to std::deque<XLCellValue>
    cout << endl << "Conversion to std::deque<XLCellValue> ..." << endl;
    for (auto& row : wks.rows()) {
        for (auto& value : std::deque<XLCellValue>(row.values())) {
            cout << value << " ";
        }
        cout << endl;
    }

    // Conversion to std::list<XLCellValue>
    cout << endl << "Conversion to std::list<XLCellValue> ..." << endl;
    for (auto& row : wks.rows()) {
        for (auto& value : std::list<XLCellValue>(row.values())) {
            cout << value << " ";
        }
        cout << endl;
    }

    // In addition to supporting any bi-directional container types, the cell values can also be converted to
    // compatible plain data types instead of XLCellValue objects, e.g. ints, doubles, bools and std::string.
    // This is illustrated in the following three blocks:

    // Conversion to std::vector<[int, double, bool, std::string]>
    cout << endl << "Conversion to std::vector<[int, double, bool, std::string]> ..." << endl;
    for (auto& value : std::vector<int>(wks.row(1).values())) cout << value << " ";
    cout << endl;
    for (auto& value : std::vector<double>(wks.row(2).values())) cout << value << " ";
    cout << endl;
    for (bool value : std::vector<bool>(wks.row(3).values())) cout << value << " ";
    cout << endl;
    for (auto& value : std::vector<std::string>(wks.row(4).values())) cout << value << " ";
    cout << endl;

    // Conversion to std::deque<[int, double, bool, std::string]>
    cout << endl << "Conversion to std::deque<[int, double, bool, std::string]> ..." << endl;
    for (auto& value : std::deque<int>(wks.row(1).values())) cout << value << " ";
    cout << endl;
    for (auto& value : std::deque<double>(wks.row(2).values())) cout << value << " ";
    cout << endl;
    for (auto& value : std::deque<bool>(wks.row(3).values())) cout << value << " ";
    cout << endl;
    for (auto& value : std::deque<std::string>(wks.row(4).values())) cout << value << " ";
    cout << endl;

    // Conversion to std::list<[int, double, bool, std::string]>
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
