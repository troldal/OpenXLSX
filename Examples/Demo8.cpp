#include <OpenXLSX.hpp>
#include <iostream>

using namespace std;
using namespace OpenXLSX;

int main()
{
    XLDocument doc;
    doc.create("demo8.xlsx");
    auto wb = doc.workbook();
    auto ws = wb.worksheet("Sheet1");

    std::vector<XLCellValue> rowValues;
    std::string testString = "TestString";
    rowValues.emplace_back(testString); //produces Excel problem
    rowValues.emplace_back(std::string("SomeString")); //produces Excel problem
    rowValues.emplace_back("char star"); //works fine
    rowValues.emplace_back(testString.c_str()); //works fine

    ws.row(1).values() = rowValues;

    doc.save();

    return 0;
}