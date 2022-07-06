#ifndef OPENXLSX_DEMO7_HPP
#define OPENXLSX_DEMO7_HPP

#include <OpenXLSX.hpp>

#include <chrono>
#include <deque>
#include <iostream>
#include <numeric>
#include <random>

int demo7()
{
    using namespace std;
    using namespace OpenXLSX;

    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #07: Row Handling (using iterators)\n";
    cout << "********************************************************************************\n";

    // As an alternative to using cell ranges, you can use row ranges and iterate through the cell values
    // This is somewhat slower than using standard containers, but may be convenient in some cases.

    using std::chrono::high_resolution_clock;
    using std::chrono::duration_cast;
    using std::chrono::duration;
    using std::chrono::milliseconds;
    using std::chrono::seconds;

    decltype(high_resolution_clock::now()) t1;
    decltype(high_resolution_clock::now()) t2;
    duration<double> dur;

    // First, create a new document and access the sheet named 'Sheet1'.
    cout << "\nGenerating spreadsheet ... ";
    cout.flush();

    t1 = high_resolution_clock::now();
    XLDocument doc;
    std::string path = "./Demo07.xlsx";
    doc.create(path);
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
    t2 = high_resolution_clock::now();
    dur = t2 - t1;
    std::cout << "(" << dur.count() << " seconds)\n";

    // Saving a large spreadsheet can take a while...
    cout << "Saving spreadsheet ... ";
    cout.flush();
    t1 = high_resolution_clock::now();
    doc.save();
    doc.close();
    t2 = high_resolution_clock::now();
    dur = t2 - t1;
    std::cout << "(" << dur.count() << " seconds)\n";

    // Reopen the spreadsheet...
    cout << "Re-opening spreadsheet ... ";
    cout.flush();
    t1 = high_resolution_clock::now();
    doc.open(path);
    wks = doc.workbook().worksheet("Sheet1");
    t2 = high_resolution_clock::now();
    dur = t2 - t1;
    std::cout << "(" << dur.count() << " seconds)\n";

    // Prepare for reading...
    uint64_t sum = 0;
    uint64_t count = 0;
    cout << "Reading data from spreadsheet ...";

    // The 'cells()' member function can be used in the usual algorithm in the
    // STL, such as count_if and accumulate.
    t1 = high_resolution_clock::now();
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

    t2 = high_resolution_clock::now();
    dur = t2 - t1;
    std::cout << " (" << dur.count() << " seconds)\n";

    cout << "Cell count: " << count << endl;
    cout << "Sum of cell values: " << sum << endl;

    doc.close();

    cout << "\nDEMO PROGRAM #07 COMPLETED\n\n";

    return 0;
}


#endif    // OPENXLSX_DEMO7_HPP
