#include <OpenXLSX.hpp>
// TODO documentation !!!!
using namespace OpenXLSX;

int main() {

    XLDocument doc;
    doc.open("./CTest.xlsx");
    auto wks = doc.workbook().worksheet("Input");


    auto rng = wks.range(XLCellReference("A1"), XLCellReference("C20"));
    XLNamedRange testrg = doc.workbook().namedRange("defined_in_sheet");

    testrg = doc.workbook().namedRange("testRange");
    XLCellValue iteration;
    for (auto& cell : testrg) 
        iteration = cell.value();

    doc.workbook().deleteNamedRange("NewFromCpp");

    XLCell i = testrg.firstCell();
    XLCellValue itet = i.value();
    i.value() = "Changed By ASH";

    std::string n = testrg.name();

    doc.workbook().addNamedRange("NewFromCpp","Input!$F$4f",2);


    auto testTable = doc.workbook().table("tbl_Hullform");


    doc.save();
    doc.close();

    return 0;
}