#include <OpenXLSX.hpp>
#include <iostream>

using namespace std;
using namespace OpenXLSX;

void printWorkbook(const XLWorkbook& wb) {
    cout << "\nSheets in workbook:\n";
    for (const auto& name : wb.worksheetNames()) cout << wb.indexOfSheet(name) << " : " << name << "\n";
}

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #03: Sheet Handling\n";
    cout << "********************************************************************************\n";

    // OpenXLSX can be used to create and delete sheets in a workbook, as well as re-ordering of sheets.
    // This example illustrates how this can be done. Please note that at the moment, chartsheets can only
    // be renamed and deleted, not created or manipulated.

    // First, create a new document and store the workbook object in a variable. Print the sheet names.
    XLDocument doc;
    doc.create("./Demo03.xlsx");
    auto wbk = doc.workbook();
    printWorkbook(wbk);

    // Add two new worksheets. The 'addWorksheet' method takes the name of the new sheet as an argument,
    // and appends the new workdheet at the end.
    // Only worksheets can be added; there is no 'addChartsheet' method.
    wbk.addWorksheet("Sheet2");
    wbk.addWorksheet("Sheet3");
    printWorkbook(wbk);

    cout << "Sheet1 active: "  << (wbk.worksheet("Sheet1").isActive() ? "true" : "false") << endl;
    cout << "Sheet2 active: "  << (wbk.worksheet("Sheet2").isActive() ? "true" : "false") << endl;
    cout << "Sheet3 active: "  << (wbk.worksheet("Sheet3").isActive() ? "true" : "false") << endl;

    wbk.worksheet("Sheet3").setActive();

    cout << "Sheet1 active: "  << (wbk.worksheet("Sheet1").isActive() ? "true" : "false") << endl;
    cout << "Sheet2 active: "  << (wbk.worksheet("Sheet2").isActive() ? "true" : "false") << endl;
    cout << "Sheet3 active: "  << (wbk.worksheet("Sheet3").isActive() ? "true" : "false") << endl;

    // OpenXLSX provides three different classes to handle workbook sheets: XLSheet, XLWorksheet, and
    // XLChartsheet. As you might have guessed, XLWorksheet and XLChartsheet represents worksheets
    // and chartsheets, respectively. XLChartsheet only has a limited set of functionality.
    // XLSheet, on the other hand, is basically a std::variant that can hold either an XLWorksheet,
    // or an XLChartsheet object. XLSheet has a limited set of functions that are common to
    // both XLWorksheet and XLChartsheet objects, such as 'clone()' or 'setIndex()'. XLSheet is
    // not a parent class to either XLWorksheet or XLChartsheet, and therefore cannot be used
    // polymorphically.

    // From an XLSheet object you can retrieve the contained XLWorksheet or XLChartsheet object by
    // using the 'get<>()' function:
    auto s1 = wbk.sheet("Sheet2").get<XLWorksheet>();

    // Alternatively, you can retrieve the contained object, by using implicit conversion:
    XLWorksheet s2 = wbk.sheet("Sheet2");

    // Existing sheets can be cloned by calling the 'clone' method on the individual sheet,
    // or by calling the 'cloneSheet' method from the XLWorkbook object. If the latter is
    // chosen, both the name of the sheet to be cloned, as well as the name of the new
    // sheet must be provided.
    // In principle, chartsheets can also be cloned, but the results may not be as expected.
    wbk.sheet("Sheet1").clone("Sheet4");
    wbk.cloneSheet("Sheet2", "Sheet5");
    printWorkbook(wbk);

    // The sheets in the workbook can be reordered by calling the 'setIndex' method on the
    // individual sheets (or worksheets/chartsheets).
    wbk.deleteSheet("Sheet1");
    wbk.worksheet("Sheet5").setIndex(1);
    wbk.worksheet("Sheet4").setIndex(2);
    wbk.worksheet("Sheet3").setIndex(3);
    wbk.worksheet("Sheet2").setIndex(4);
    printWorkbook(wbk);

    // The color of each sheet tab can be set using the 'setColor' method for a
    // sheet, and passing an XLColor object as an argument.
    wbk.sheet("Sheet2").setColor(XLColor(0, 0, 0));
    wbk.sheet("Sheet3").setColor(XLColor(255, 0, 0));
    wbk.sheet("Sheet4").setColor(XLColor(0, 255, 0));
    wbk.sheet("Sheet5").setColor(XLColor(0, 0, 255));

    doc.save();

    return 0;
}
