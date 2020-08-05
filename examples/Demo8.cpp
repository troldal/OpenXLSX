#include <OpenXLSX.hpp>

using namespace std;
using namespace OpenXLSX;

int main()
{
    XLDocument doc;
    doc.create("./MyTest.xlsx");
    auto wks = doc.workbook().worksheet("Sheet1");

    auto rng = wks.range(XLCellReference("A1"), XLCellReference(1024, 1024));

    for (auto iter = rng.begin(); iter != rng.end(); ++iter) iter->value() = "OpenXLSX";

    doc.save();

    return 0;
}
