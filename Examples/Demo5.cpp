#pragma warning(push)
#pragma warning(disable : 4244)

#include <OpenXLSX.hpp>
#include <iostream>
#include <numeric>
#include <random>

using namespace std;
using namespace OpenXLSX;

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #05: Ranges and Iterators\n";
    cout << "********************************************************************************\n";

    // With OpenXLSX, a range of cells can be defined, in order to iterate through the
    // cells in the range. A range is a quardratic region of cells in a worksheet.
    // A range can be defined using the 'range()' method on an XLWorksheet object. This function
    // takes two XLCellReferences: the cell in the upper left corner of the range, and the cell
    // in the lower right corner.
    // The resulting XLCellRange objects provides 'begin()' and 'end()' (returning iterators),
    // so that the object can be used in a range-based for-loop.

    // First, create a new document and access the sheet named 'Sheet1'.
    cout << "\nGenerating spreadsheet ..." << endl;
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

    // Saving a large spreadsheet can take a while...
    cout << "Saving spreadsheet ..." << endl;
    doc.save();
    doc.close();

    // Reopen the spreadsheet..
    cout << "Re-opening spreadsheet ..." << endl;
    doc.open("./Demo05.xlsx");
    wks = doc.workbook().worksheet("Sheet1");

    // Create the range, count the cells in the range, and sum the numbers assigned
    // to the cells in the range.
    cout << "Reading data from spreadsheet ..." << endl;
    rng = wks.range(XLCellReference("A1"), XLCellReference(OpenXLSX::MAX_ROWS, 8));
    cout << "Cell count: " << std::distance(rng.begin(), rng.end()) << endl;
    cout << "Sum of cell values: "
         << accumulate(rng.begin(), rng.end(), 0, [](uint64_t a, XLCell& b) { return a + b.value().get<int64_t>(); });

    doc.close();

    return 0;
}

#pragma warning(pop)
