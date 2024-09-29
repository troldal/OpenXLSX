#include <OpenXLSX.hpp>
#include <iostream>
#include <random>
#include <deque>
#include <numeric>

using namespace std;
using namespace OpenXLSX;

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #07: Row Handling (using iterators)\n";
    cout << "********************************************************************************\n";

    // As an alternative to using cell ranges, you can use row ranges.
    // This can be significantly faster (up to twice as fast for both reading and
    // writing).

    // First, create a new document and access the sheet named 'Sheet1'.
    cout << "\nGenerating spreadsheet ..." << endl;
    XLDocument doc;
    std::string path = "./Demo07.xlsx";
    doc.create(path, XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // Create a random-number generator (to be used later)
    std::random_device                 rand_dev;
    std::mt19937                       generator(rand_dev());
    std::uniform_int_distribution<int> distr(0, 99);

    // The XLWorksheet class has a 'rows()' method, that returns a XLRowRange
    // object. Similar to XLCellRange, XLRowRange provides begin and end iterators,
    // enabeling iteration through the individual XLRow objects.
    for (auto& row : wks.rows(OpenXLSX::MAX_ROWS)) {
        for (auto cell: row.cells(8)) cell.value() = distr(generator);

        // The XLRow class provides a 'cells()' method. This provides begin and end
        // iterators to the cells, or range of cells, in the row. Iterating through
        // the cells works in the usual manner and values can be read and written.
    }

    // Saving a large spreadsheet can take a while...
    cout << "Saving spreadsheet ..." << endl;
    doc.save();
    doc.close();

    // Reopen the spreadsheet...
    cout << "Re-opening spreadsheet ..." << endl;
    doc.open(path);
    wks = doc.workbook().worksheet("Sheet1");

    // Prepare for reading...
    uint64_t sum = 0;
    uint64_t count = 0;
    cout << "Reading data from spreadsheet ..." << endl;

    // The 'cells()' member function can be used in the usual algorithm in the
    // STL, such as count_if and accumulate.
    for (auto& row : wks.rows()) {
        // Count the number of cell values
        count += std::count_if(row.cells().begin(), row.cells().end(), [](const XLCell& c) {
            return c.value().type() != XLValueType::Empty;
        });

        // Sum the numbers in each cell.
        sum += std::accumulate(row.cells().begin(),
                               row.cells().end(),
                               0,
                               [](uint64_t a, XLCell& b) {return a + b.value().get<uint64_t>(); });
    }

    cout << "Cell count: " << count << endl;
    cout << "Sum of cell values: " << sum << endl;

    doc.close();

    return 0;
}
