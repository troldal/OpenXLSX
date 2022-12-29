#include <OpenXLSX.hpp>

using namespace OpenXLSX;

int main() {

    XLDocument doc;
    doc.open("./testSimple.xlsx");
    auto wks = doc.workbook().worksheet("Matrix");

    auto matrix = wks.range("C4:G16");
    int myCell = matrix[57].value();


    auto testTable = doc.workbook().table("tbl_one");


    //testTable.setName("tbl_ash");
    auto testName = testTable.name();
    auto cols = testTable.columnNames();

    auto ncol = testTable.columnIndex("Col bord");
    auto sht = testTable.getWorksheet();
    auto rang = testTable.dataBodyRange();
    auto rows = testTable.tableRows();
    
    int val = 67;
    for(auto& row:rows) {
        row[0].value() = val;
        val +=1;
    }
/*
    for(auto& cell:rows[9]) {
        int i=cell.value();
    }
    */

    auto row5 = rows[5];


    std::string shtName = sht.name();

    //doc.workbook().definedName("nbAngles");

    //wks.cell("A1").value() = "Hello, OpenXLSX ASHHHH!";

    doc.save();
    doc.close();

    return 0;
}