#include <OpenXLSX.hpp>
#include <iostream>
#include <random>
#include <deque>

using namespace std;
using namespace OpenXLSX;

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #06: Row Handling\n";
    cout << "********************************************************************************\n";

    cout << "\nGenerating spreadsheet (1,048,576 rows x 8 columns) ..." << endl;
    XLDocument doc;
    doc.create("./Demo06.xlsx");
    auto wks = doc.workbook().worksheet("Sheet1");

    std::random_device                 rand_dev;
    std::mt19937                       generator(rand_dev());
    std::uniform_int_distribution<int> distr(0, 99);

    std::vector<XLCellValue> writeValues;

//    for (auto& row : wks.rows(1'048'576)) {
    for (auto& row : wks.rows(10)) {
        writeValues.clear();
        writeValues.emplace_back(distr(generator));
        writeValues.emplace_back(distr(generator));
        writeValues.emplace_back(distr(generator));
        writeValues.emplace_back(distr(generator));
        writeValues.emplace_back(distr(generator));
        writeValues.emplace_back(distr(generator));
        writeValues.emplace_back(distr(generator));
        writeValues.emplace_back(distr(generator));

        row.values() = writeValues;
    }

    cout << "Saving spreadsheet (1,048,576 rows x 8 columns) ..." << endl;
    doc.save();
    doc.close();

    cout << "Re-opening spreadsheet (1,048,576 rows x 8 columns) ..." << endl;
    doc.open("./Demo06.xlsx");
    wks = doc.workbook().worksheet("Sheet1");
    uint64_t sum = 0;
    uint64_t count = 0;
//    std::vector<XLCellValue> readValues;

    cout << "Reading data from spreadsheet (1,048,576 rows x 8 columns) ..." << endl;

    for (auto& row : wks.rows()) {
        auto readValues = row.values<std::vector<XLCellValue>>();
        count += std::count_if(readValues.begin(), readValues.end(), [](const XLCellValue& v) {
               return v.type() != XLValueType::Empty;
        });
    }
    cout << "Cell count: " << count << endl;

    for (auto& row : wks.rows()) {
        auto readValues = row.values<std::vector<XLCellValue>>();
        sum += std::accumulate(readValues.begin(),
                               readValues.end(),
                                  0,
                                  [](uint64_t a, XLCellValue& b) { return a + b.get<uint64_t>(); });
    }
    cout << "Sum of cell values: " << sum << endl;

    doc.close();

    return 0;
}