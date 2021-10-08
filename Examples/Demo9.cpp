#include <OpenXLSX.hpp>
#include <deque>
#include <iostream>
#include <list>
#include <random>

using namespace std;
using namespace OpenXLSX;

int main()
{

    XLDocument doc;
    doc.open("./ress.xlsx");
    auto wks = doc.workbook().worksheet("Sheet1");

    std::vector<std::string> strings;

    for (auto row : wks.rows()){
        std::vector<XLCellValue> values = row.values();
        strings.emplace_back(values[0]);
    }

    int index = 1;
    for (auto str : strings) {
        wks.cell(index, 3).value() = str;
        ++index;
    }

    doc.saveAs("./Demo9.xlsx");
    doc.close();

    return 0;
}
