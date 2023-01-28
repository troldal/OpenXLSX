#include <OpenXLSX.hpp>

using namespace std;
using namespace OpenXLSX;

int main() {
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #09: Named Range\n";
    cout << "********************************************************************************\n";

    // This example illustrate the use of the named range object. This could be particulary
    // usefull when programming without GUI, to set up name in the Excel file, so that 
    // data could be safetly retrived and written by the code in dedicated location in the Excel file
    // Unlike other implementation, named range is accessible within workbook object. Sheet limited 
    // maed range are as well accessed via the wrokbook.

    // First, create a new document and access the 1st sheet.
    cout << "\nGenerating spreadsheet ..." << endl;
    XLDocument doc;
    doc.create("./Demo09.xlsx");
    auto wks = doc.workbook().sheet(1).get<XLWorksheet>();
    string sheetName = wks.name();

    // Define a namedrange "New name"
    // At this stage the range shall be continuous
    // a third argument could be provided the index of the spreasheet where 
    // the named range will be defined. Absolute marks ($) will be added.
    string fullRef = sheetName + "!A5:E10";
    XLNamedRange myRange = doc.workbook().addNamedRange("MyNamedRange",fullRef /*, localSheetId*/);

    // Save the sheet...
    cout << "Saving spreadsheet ..." << endl;
    doc.save();
    doc.close();

    // ...and reopen it (just to make sure that it is a valid .xlsx file)
    cout << "Re-opening spreadsheet ..." << endl << endl;
    doc.open("./Demo09.xlsx");
    wks = doc.workbook().worksheet(sheetName);

    // The NamedRange object is a range, it can be iterated
    cout << "Filling the named range object ..." << endl;
    XLNamedRange rng = doc.workbook().namedRange("MyNamedRange");
    int i = 0;
    for (auto& cell : rng){
        cell.value() = i;
        i++;
    }

    // It can also be accessed via []
    cout << "Value of the 10th cell of the named range:" << endl;
    cout << rng[10].value() << endl;

    // for convenience, in particular in case of mono cell named range, 
    // a fonction is provided: firstCell()
    cout << "first Cell Value: " << rng.firstCell().value()<< endl;

    // Finally named range could be removed
    //doc.workbook().deleteNamedRange("MyNamedRange");

    doc.save();
    doc.close();

    return 0;
}