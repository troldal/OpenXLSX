#include <OpenXLSX.hpp>

#include <vector>

using namespace std;
using namespace OpenXLSX;

int main() {
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #10: Tables first part\n";
    cout << "********************************************************************************\n";

    // This example illustrate the use of the named range object. This could be particulary
    // usefull when programming without GUI, to set up name in the Excel file, so that 
    // data could be safetly retrived and written by the code in dedicated location in the Excel file
    // Unlike other implementation, named range is accessible within workbook object. Sheet limited 
    // maed range are as well accessed via the wrokbook.

    // First, create a new document and access the 1st sheet.
    cout << "\nGenerating spreadsheet ..." << endl;
    XLDocument doc;
    doc.create("./Demo10.xlsx");
    auto wks = doc.workbook().sheet(1).get<XLWorksheet>();
    string sheetName = wks.name();

    // Fill the column header and create the table
    // If the table header is not set, default name will be setup
    // Also columns name shall be unique, a increment marks is added if required
    cout << "Creating table ..." << endl;
    wks.cell("B2").value() = "#";
    wks.cell("C2").value() = "One";
    wks.cell("D2").value() = "Col";
    wks.cell("E2").value() = "With";
    wks.cell("F2").value() = "Table";

    XLTable myTable = doc.workbook().addTable(sheetName,"MyTable","B2:F18");

    // Save the sheet...
    cout << "Saving spreadsheet ..." << endl;
    doc.save();
    doc.close();

    // ...and reopen it (just to make sure that it is a valid .xlsx file)
    cout << "Re-opening spreadsheet ..." << endl << endl;
    doc.open("./Demo10.xlsx");
    wks = doc.workbook().worksheet(sheetName);

    // We can retrieve the headers name:
    XLTable tbl = doc.workbook().table("MyTable");
    vector<std::string> headers = tbl.columnNames();

    cout << "Table "<< tbl.name() << " header names:" << endl;
    int i = 0;
    for(const auto& h : headers){
        cout << i++ << " - " << h << endl;
    }

    // We can also access different items
    auto ncol       = tbl.columnIndex("Col"); // Retrieve the colum index by the name
    auto sht        = tbl.getWorksheet();      // Retrieve and object worksheet
    auto rang       = tbl.tableRange();       // whole table range including total & headers
    auto bodyRange  = tbl.dataBodyRange();    // Only data body table range (excl. total & headers)
    auto nRow       = tbl.rowsCount();    // Only data body table range (excl. total & headers)
    // And iterate throught the table rows
    auto nCol       = tbl.columnsCount();    // Only data body table range (excl. total & headers)
    
    int j = 0;
    for(const auto& row : tbl.tableRows()){
        row[tbl.columnIndex("#")].value() = j;
        row[1].value() = "Data" + to_string(j);
        row[tbl.columnIndex("Table")].value() = (float)j * 2.0f / 3.0f;
        j++;
    }

    // loop could also be done on colums
    for(auto& col : tbl.columns())
        for (auto& cell : col.bodyRange())
            cout << cell.value() << " - ";
    
    cout << endl;

    // Also show the total with selected function
    tbl.autofilter().hideArrows();
    //tbl.setHeaderVisible(false);

    // Total formulas 
    tbl.setTotalVisible(true);
    //tbl.column("Table").setTotalsRowFunction("");
    //tbl.column("Table").setTotalsRowFunction("count");
    tbl.column("Table").totalsRowFormula() ="sum";
    string totaFormula = tbl.column("Table").totalsRowFormula();
    cout << "total Formula in the table column : " << totaFormula << endl;

    // To clear, a empty string could be sent, or the method 
    // clearTotalsRowFormula could be called
    tbl.column("Table").totalsRowFormula() ="";

    // Columns formulas could be setup, check and cleared, using either
    // empty string or calling clearColumnFormula
    tbl.column("With").columnFormula() = "MyTable[[#This Row],['#]]*2";
    string columFormula =  tbl.column("With").columnFormula();
    cout << "Column Formula : " << columFormula << endl;


    // Table style basics
    cout << "Table Style : " << tbl.tableStyle().style() << endl;
    tbl.tableStyle().setStyle("TableStyleDark7"); 
    
    cout << "Column stripes : " << tbl.tableStyle().columnStripes() << endl;
    cout << "row stripes : " << tbl.tableStyle().rowStripes() << endl;

    tbl.tableStyle().showRowStripes(false);
    tbl.tableStyle().showColumnStripes(true);

    cout << "1st Column highlighted : " << tbl.tableStyle().firstColumnHighlighted() << endl;
    cout << "last Column highlighted : " << tbl.tableStyle().lastColumnHighlighted() << endl;

    tbl.tableStyle().showFirstColumnHighlighted(true);
    tbl.tableStyle().showLastColumnHighlighted(true);

    doc.save();
    doc.close();

    return 0;
}