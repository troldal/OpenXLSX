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
         << accumulate(rng.begin(), rng.end(), 0, [](uint64_t a, XLCell& b) { return a + b.value().get<uint64_t>(); });

    doc.close();

    return 0;
}