#ifndef OPENXLSX_DEMO5_HPP
#define OPENXLSX_DEMO5_HPP

#pragma warning(push)
#pragma warning(disable : 4244)

#include <OpenXLSX.hpp>

#include <chrono>
#include <iostream>
#include <numeric>
#include <random>

int demo5()
{
    using namespace std;
    using namespace OpenXLSX;

    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #05: Ranges and Iterators\n";
    cout << "********************************************************************************\n";

    // With OpenXLSX, a range of cells can be defined, in order to iterate through the
    // cells in the range. A range is a quadratic region of cells in a worksheet.
    // A range can be defined using the 'range()' method on an XLWorksheet object. This function
    // takes two XLCellReferences: the cell in the upper left corner of the range, and the cell
    // in the lower right corner.
    // The resulting XLCellRange objects provides 'begin()' and 'end()' (returning iterators),
    // so that the object can be used in a range-based for-loop.

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
    doc.create("./Demo05.xlsx");
    auto wks = doc.workbook().worksheet("Sheet1");

    // Create a random-number generator (to be used later)
    std::random_device                 rand_dev;
    std::mt19937                       generator(rand_dev());
    std::uniform_int_distribution<int> distr(0, 99);

    // Create a cell range using the 'range()' method of the XLWorksheet class.
    // The 'range()' method takes two XLCellReference objects; one for the cell
    // in the upper left corner, and one for the cell in the lower right corner.
    // OpenXLSX defines two constants: MAX_ROWS which is the maximum number of
    // rows in a worksheet, and MAX_COLS which is the maximum number of columns.
    auto rng = wks.range(XLCellReference("A1"), XLCellReference(OpenXLSX::MAX_ROWS, 8));

    // An XLCellRange object provides begin and end iterators, and can be used in
    // range-based for-loops. The next line iterates through all the cells in the
    // range and assigns a random number to each cell:
    for (auto& cell : rng) cell.value() = distr(generator);
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

    // Reopen the spreadsheet..
    cout << "Re-opening spreadsheet ... ";
    cout.flush();
    t1 = high_resolution_clock::now();
    doc.open("./Demo05.xlsx");
    wks = doc.workbook().worksheet("Sheet1");
    t2 = high_resolution_clock::now();
    dur = t2 - t1;
    std::cout << "(" << dur.count() << " seconds)\n";

    // Create the range, count the cells in the range, and sum the numbers assigned
    // to the cells in the range.
    uint64_t sum = 0;
    uint64_t count = 0;
    cout << "Reading data from spreadsheet ...";

    t1 = high_resolution_clock::now();
    rng = wks.range(XLCellReference("A1"), XLCellReference(OpenXLSX::MAX_ROWS, 8));
    sum = accumulate(rng.begin(), rng.end(), 0, [&](uint64_t a, XLCell& b) { count++; return a + b.value().get<int64_t>(); });
    t2 = high_resolution_clock::now();
    dur = t2 - t1;
    std::cout << " (" << dur.count() << " seconds)\n";

    cout << "Cell count: " << count << endl; //std::distance(rng.begin(), rng.end());
    cout << "Sum of cell values: " << sum << endl;

    doc.close();

    cout << "\nDEMO PROGRAM #05 COMPLETED\n\n";

    return 0;
}

#pragma warning(pop)

#endif    // OPENXLSX_DEMO5_HPP
