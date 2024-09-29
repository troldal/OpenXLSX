  
#include <OpenXLSX.hpp>
#include <iostream>

using namespace std;
using namespace OpenXLSX;

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #02: Formulas\n";
    cout << "********************************************************************************\n";

    XLDocument doc;
    doc.create("./Demo02.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // Similar cell values, which are represented by XLCellValue objects, formulas are represented
    // by XLFormula objects. They can be accessed through the XLCell interface using the .formula()
    // member function. It should be noted, however, that the functionality of XLFormula is somewhat
    // limited. Excel often uses 'shared' formulas, where the same formula is applied to several
    // cells. XLFormula cannot handle shared formulas. Also, it cannot handle array formulas. This,
    // in effect, means that XLFormula is not very useful for reading formulas from existing spread-
    // sheets, but should rather be used to add or overwrite formulas to spreadsheets.

    wks.cell("A1").value() = "Number:";
    wks.cell("B1").value() = 1;
    wks.cell("C1").value() = 2;
    wks.cell("D1").value() = 3;

    // Formulas can be added to a cell using the .formula() method on an XLCell object. The formula can
    // be added by creating a separate XLFormula object, or (as shown below) by assigning a string
    // holding the formula;
    // Nota that OpenXLSX does not calculate the results of a formula. If you add a formula
    // to a spreadsheet using OpenXLSX, you have to open the spreadsheet in the Excel application
    // in order to see the results of the calculation.
    wks.cell("A2").value() = "Calculation:";
    wks.cell("B2").formula() = "SQRT(B1)";
    wks.cell("C2").formula() = "EXP(C1)";
    wks.cell("D2").formula() = "LN(D1)";

    // The XLFormula object can be retrieved using the .formula() member function.
    // The XLFormula .get() method can be used to get the formula string.
    XLFormula B2 = wks.cell("B2").formula();
    XLFormula C2 = wks.cell("C2").formula();
    XLFormula D2 = wks.cell("D2").formula();

    cout << "Formula in B2: " << B2.get() << endl;
    cout << "Formula in C2: " << C2.get() << endl;
    cout << "Formula in D2: " << D2.get() << endl << endl;

    // Alternatively, the XLFormula object can be implicitly converted to a string object.
    std::string B2string = wks.cell("B2").formula();
    std::string C2string = wks.cell("C2").formula();
    std::string D2string = wks.cell("D2").formula();

    cout << "Formula in B2: " << B2string << endl;
    cout << "Formula in C2: " << C2string << endl;
    cout << "Formula in D2: " << D2string << endl << endl;

    doc.save();
    doc.close();

    return 0;
}
