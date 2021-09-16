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

    cout << "\nGenerating spreadsheet (1,048,576 rows x 8 columns) ..." << endl;
    XLDocument doc;
    doc.create("./Demo05.xlsx");
    auto wks = doc.workbook().worksheet("Sheet1");
    auto rng = wks.range(XLCellReference("A1"), XLCellReference(1048576, 8));

    std::random_device                 rand_dev;
    std::mt19937                       generator(rand_dev());
    std::uniform_int_distribution<int> distr(0, 99);

    for (auto& cell : rng) cell.value() = distr(generator);

    cout << "Saving spreadsheet (1,048,576 rows x 8 columns) ..." << endl;
    doc.save();
    doc.close();

    cout << "Re-opening spreadsheet (1,048,576 rows x 8 columns) ..." << endl;
    doc.open("./Demo05.xlsx");
    wks = doc.workbook().worksheet("Sheet1");
    rng = wks.range(XLCellReference("A1"), XLCellReference(1048576, 8));

    cout << "Reading data from spreadsheet (1,048,576 rows x 8 columns) ..." << endl;
    cout << "Cell count: " << std::distance(rng.begin(), rng.end()) << endl;
    cout << "Sum of cell values: "
         << accumulate(rng.begin(), rng.end(), 0, [](uint64_t a, XLCell& b) { return a + b.value().get<int64_t>(); });

    doc.close();

    return 0;
}

#pragma warning(pop)
