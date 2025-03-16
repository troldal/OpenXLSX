#include <OpenXLSX.hpp>
#include <iostream>
#include <cmath>

using namespace std;
using namespace OpenXLSX;

/**
 * @brief demo how to iterate over all worksheet comments by sequence in XML - for few comments, this is substantially faster than testing each cell for a non-empty comment
 * @param doc the XLDocument
 * @return n/a
 */
void printAllDocumentComments(XLDocument const & doc)
{
    for( size_t i = 1; i <= doc.workbook().worksheetCount(); ++i ) {
        auto wks = doc.workbook().worksheet(i);
        if( wks.hasComments() ) {
            std::cout << "worksheet(" << i << ") with name \"" << wks.name() << "\" has comments" << std::endl;
            XLComments wksComments = wks.comments();
            size_t commentCount = wksComments.count();
            for( size_t idx = 0; idx < commentCount; ++idx ) {
                XLComment com = wksComments.get(idx);
                std::cout << "comment in " << com.ref() << ": \"" << com.text() << "\" (author: " << wksComments.author(com.authorId()) << ")" << std::endl;
            }
        }
    }
}

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #01: Basic Usage and XLWorksheet protection\n";
    cout << "********************************************************************************\n";

    // This example program illustrates basic usage of OpenXLSX, for example creation of a new workbook, and read/write
    // of cell values.

    // First, create a new document and access the sheet named 'Sheet1'.
    // New documents contain a single worksheet named 'Sheet1'
    XLDocument doc;
    doc.create("./Demo01.xlsx", XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");

    // The individual cells can be accessed by using the .cell() method on the worksheet object.
    // The .cell() method can take the cell address as a string, or alternatively take a XLCellReference
    // object. By using an XLCellReference object, the cells can be accessed by row/column coordinates.
    // The .cell() method returns an XLCell object.

    // The .value() method of an XLCell object can be used for both getting and setting the cell value.
    // Setting the value of a cell can be done by using the assignment operator on the .value() method
    // as shown below. Alternatively, a .set() can be used. The cell values can be floating point numbers,
    // integers, strings, and booleans. It can also accept XLDateTime objects, but this requires special
    // handling (see later).
    wks.cell("A1").value() = 3.14159265358979323846;
    wks.cell("B1").value() = 42;
    wks.cell("C1").value() = "  Hello OpenXLSX!  ";
    wks.cell("D1").value() = true;
    wks.cell("E1").value() = std::sqrt(-2); // Result is NAN, resulting in an error value in the Excel spreadsheet.

    // As mentioned, the .value() method can also be used for getting tha value of a cell.
    // The .value() method returns a proxy object that cannot be copied or assigned, but
    // it can be implicitly converted to an XLCellValue object, as shown below.
    // Unfortunately, it is not possible to use the 'auto' keyword, so the XLCellValue
    // type has to be explicitly stated.
    XLCellValue A1 = wks.cell("A1").value();
    XLCellValue B1 = wks.cell("B1").value();
    XLCellValue C1 = wks.cell("C1").value();
    XLCellValue D1 = wks.cell("D1").value();
    XLCellValue E1 = wks.cell("E1").value();

    // The cell value can be implicitly converted to a basic c++ type. However, if the type does not
    // match the type contained in the XLCellValue object (if, for example, floating point value is
    // assigned to a std::string), then an XLValueTypeError exception will be thrown.
    // To check which type is contained, use the .type() method, which will return a XLValueType enum
    // representing the type. As a convenience, the .typeAsString() method returns the type as a string,
    // which can be useful when printing to console.
    double vA1 = wks.cell("A1").value();
    int vB1 = wks.cell("B1").value();
    std::string vC1 = wks.cell("C1").value();
    bool vD1 = wks.cell("D1").value();
    double vE1 = wks.cell("E1").value();

    cout << "Cell A1: (" << A1.typeAsString() << ") " << vA1 << endl;
    cout << "Cell B1: (" << B1.typeAsString() << ") " << vB1 << endl;
    cout << "Cell C1: (" << C1.typeAsString() << ") " << vC1 << endl;
    cout << "Cell D1: (" << D1.typeAsString() << ") " << vD1 << endl;
    cout << "Cell E1: (" << E1.typeAsString() << ") " << vE1 << endl << endl;

    // Instead of using implicit (or explicit) conversion, the underlying value can also be retrieved
    // using the .get() method. This is a templated member function, which takes the desired type
    // as a template argument.
    cout << "Cell A1: (" << A1.typeAsString() << ") " << A1.get<double>() << endl;
    cout << "Cell B1: (" << B1.typeAsString() << ") " << B1.get<int64_t>() << endl;
    cout << "Cell C1: (" << C1.typeAsString() << ") " << C1.get<std::string>() << endl;
    cout << "Cell D1: (" << D1.typeAsString() << ") " << D1.get<bool>() << endl;
    cout << "Cell E1: (" << E1.typeAsString() << ") " << E1.get<double>() << endl << endl;

    // XLCellValue objects can also be copied and assigned to other cells. This following line
    // will copy and assign the value of cell C1 to cell E1. Note that only the value is copied;
    // other cell properties of the target cell remain unchanged.

    wks.cell("F1").value() = wks.cell(XLCellReference("C1")).value();

    XLCellValue F1 = wks.cell("F1").value();
    cout << "Cell F1: (" << F1.typeAsString() << ") " << F1.get<std::string>() << endl << endl;

    // Date/time values is a special case. In Excel, date/time values are essentially just a
    // 64-bit floating point value, that is rendered as a date/time string using special
    // formatting. When retrieving the cell value, it is just a floating point value,
    // and there is no way to identify it as a date/time value.
    // If, however, you know it to be a date time value, or if you want to assign a date/time
    // value to a cell, you can use the XLDateTime class, which falilitates conversion between
    // Excel date/time serial numbers, and the std::tm struct, that is used to store
    // date/time data. See https://en.cppreference.com/w/cpp/chrono/c/tm for more information.

    // An XLDateTime object can be created from a std::tm object:
    std::tm tm;
    tm.tm_year = 121;
    tm.tm_mon = 8;
    tm.tm_mday = 1;
    tm.tm_hour = 12;
    tm.tm_min = 0;
    tm.tm_sec = 0;
    XLDateTime dt (tm);
//    XLDateTime dt (43791.583333333299);

    // The std::tm object can be assigned to a cell value in the same way as shown previously.
    wks.cell("G1").value() = dt;

    // And as seen previously, an XLCellValue object can be retrieved. However, the object
    // will just contain a floating point value; there is no way to identify it as a date/time value.
    XLCellValue G1 = wks.cell("G1").value();
    cout << "Cell G1: (" << G1.typeAsString() << ") " << G1.get<double>() << endl;

    // If it is known to be a date/time value, the cell value can be converted to an XLDateTime object.
    auto result = G1.get<XLDateTime>();

    // The Excel date/time serial number can be retrieved using the .serial() method.
    cout << "Cell G1: (" << G1.typeAsString() << ") " << result.serial() << endl;

    // Using the .tm() method, the corresponding std::tm object can be retrieved.
    auto tmo = result.tm();
    cout << "Cell G1: (" << G1.typeAsString() << ") " << std::asctime(&tmo);

    std::cout << std::endl;
    std::cout << "XLWorksheet comments demo" << std::endl;
    std::cout << "=========================" << std::endl;
    std::cout <<   "wks.hasComments()   is " << (wks.hasComments()   ? " TRUE" : "FALSE")
              << ", wks.hasVmlDrawing() is " << (wks.hasVmlDrawing() ? " TRUE" : "FALSE")
              << ", wks.hasTables()     is " << (wks.hasTables()     ? " TRUE" : "FALSE") << std::endl;

    XLComments &wksComments = wks.comments();   // fetch comments (and create if missing)
    std::cout << "wks.comments()      is " << (wksComments.valid() ? "valid" : "not valid") << std::endl;
    XLTables &wksTables = wks.tables();         // fetch tables (and create if missing)
    std::cout << "wks.tables()        is " << (wksTables.valid()   ? "valid" : "not valid") << std::endl;

    std::cout <<   "wks.hasComments()   is " << (wks.hasComments()   ? " TRUE" : "FALSE")
              << ", wks.hasVmlDrawing() is " << (wks.hasVmlDrawing() ? " TRUE" : "FALSE")
              << ", wks.hasTables()     is " << (wks.hasTables()     ? " TRUE" : "FALSE") << std::endl;

    // 2025-01-17: LibreOffice displays all comments with the same, LibreOffice, author - regardless of author id - to be tested with MS Office
    std::string authors[10] { "Edgar Allan Poe", "Agatha Christie", "J.R.R. Tolkien", "David Eddings", "Daniel Suarez",
    /**/                      "Mary Shelley", "George Orwell", "Stanislaw Lem", "Ray Bradbury", "William Shakespeare" };

    for( std::string author : authors ) {
        wksComments.addAuthor( author );
        uint16_t authorCount = wksComments.authorCount();
        std::cout << "wksComments.authorCount is " << authorCount << std::endl;
        std::cout << "wksComments.author(" << (authorCount - 1) << ") is " << wksComments.author(authorCount - 1) << std::endl;
    }

    wksComments.deleteAuthor(7);
    wksComments.deleteAuthor(5);
    wksComments.deleteAuthor(3);
    uint16_t authorCount = wksComments.authorCount();
    std::cout << "wksComments.authorCount is " << authorCount << " after deleting authors 7, 5 and 3" << std::endl;
    for( uint16_t index = 0; index < authorCount; ++index )
        std::cout << "wksComments.author(" << (index) << ") is " << wksComments.author(index) << std::endl;

    wksComments.set("A1", "this comment is for author #0", 0);
    wksComments.set("B2", "this comment is for author #1", 1);
    wksComments.shape("B2").style().show(); // bit cumbersome to access, but it reflects the XML: <v:shape style="[..];visibility=visible">[..]</v:shape>
    wksComments.set("C3", "this comment is for author #2", 2);
    wksComments.set("C3", "this is an updated comment for author #2", 2); // overwrite a comment
    wksComments.shape("C3").style().show();
    wksComments.set("D1", "this comment is for author #4", 4);
    wksComments.set("E2", "this comment is also for author #4", 4);
    bool deleteSuccess = wksComments.deleteComment("E2");
    std::cout << "Deleting the comment in cell E2 " << ( deleteSuccess ? "succeeded" : "did not succeed" ) << "." << std::endl;
    std::cout << std::endl;
    std::cout << "the comment in cell C4 is \"" << wksComments.get("C4") << "\"" << std::endl;
    std::cout << "the comment in cell B2 is \"" << wksComments.get("B2") << "\"" << std::endl;
    std::cout << std::endl;
    printAllDocumentComments(doc);

    std::cout << std::endl;
    std::cout << "XLWorksheet protection settings demo" << std::endl;
    std::cout << "====================================" << std::endl;
    std::cout << "Default worksheet " << wks.name() << " protection settings: " << std::endl
    /**/      << wks.sheetProtectionSummary() << std::endl;
    std::string password = "test";
    std::string pwHash = ExcelPasswordHashAsString(password);
    wks.setPassword(password);            // set password directly
    wks.setPasswordHash(pwHash);          // set password via hash
    wks.protectSheet();                   // protects all cells that have protected=true
    // protectObjects(bool set = false);     // not sure what this would do
    // protectScenarios(bool set = false);   // not sure what this would do
    wks.allowSelectUnlockedCells();       // user can select unlocked cells (default)
    wks.denySelectLockedCells();          // user can not select any cells except those that are explicitly unlocked
    // wks.denySelectUnlockedCells();        // somehow this setting REALLY does not make any sense
    // wks.clearPassword();                  // clear a password - this will not(!) disable sheet protection but will no longer require a password to unprotect the sheet in a GUI
    // wks.protectSheet(false);              // remove sheet protection - this will disable the other protect settings
    // wks.clearSheetProtection();           // this will ALL protection settings, including the password hash
    std::cout << "---" << std::endl;
    std::cout << "Configured worksheet " << wks.name() << " protection settings: " << std::endl
    /**/      << wks.sheetProtectionSummary() << std::endl;

    // Create a new cell format that unlocks cells
    XLCellFormats & cellFormats = doc.styles().cellFormats();
    XLStyleIndex newCellFormatIndex = cellFormats.create( cellFormats[ wks.cell("C1").cellFormat() ] );
    // cellFormats[ newCellFormatIndex ].setApplyProtection( true ); // seems to be irrelevant
    // cellFormats[ newCellFormatIndex ].setLocked( true );  // locked seems to be default behavior if attribute is not set
    cellFormats[ newCellFormatIndex ].setLocked( false ); // unprotect a cell

    // wks.cell("D1").setCellFormat(newCellFormatIndex);
    wks.range("C1:D1").setFormat(newCellFormatIndex); // unlock two cells
    std::cout << "====================================" << std::endl;
    std::cout << "==> the worksheet " << wks.name() << " is now protected with the password " << password << std::endl;
    std::cout << std::endl;

    std::cout << std::endl;
    std::cout << "XLWorksheet testing for empty cells demo.." << std::endl;
    std::cout << "==========================================" << std::endl;
    std::cout << "  using an XLCellIterator (recommended)" << std::endl;
    std::cout << "  -----------------------------------" << std::endl;
    XLCellRange cellRange = wks.range("A1:G2");
    for (XLCellIterator cellIt = cellRange.begin(); cellIt != cellRange.end(); ++cellIt) {
        if (!cellIt.cellExists()) continue; // prevent cell creation by access for non-existing cells
        std::cout << "cell " << cellIt.address() << " exists!" << std::endl;
    }

    std::cout << std::endl;
    std::cout << "  using XLSheet::findCell (lazy method, not recommended)" << std::endl;
    std::cout << "  NOTE: this method should only be used for tests of individual cells when performance is not an issue" << std::endl;
    std::cout << "  ------------------------------------------------------" << std::endl;
    wks.cell("D2").formula().set("=A1+B1");
    wks.cell("E2").formula().set(""); // setting an empty (zero-length) formula string will delete the formula node

    for (int row = 1; row < 4; ++row) {
        for (int col = 1; col < 8; ++col) {
            XLCellReference ref(row, col);
            XLCellAssignable c = wks.findCell(ref);
            if (c.empty())
        continue;

            // It is up to the implementation to decide which tests shall be performed on a cell that exists - here are a few exemplary tests
            std::string valueAsString = c.getString();
            if (valueAsString == "" && !c.hasFormula())
                cout << "cell " << ref.address() << " exists but has no value or formula" << std::endl;
            else {
                // cout << "cell " << ref.address() << " value is |" << c.value() << "|" << std::endl;
                if (valueAsString.length() > 0)
                    cout << "cell " << ref.address() << " value is |" << valueAsString << "|" << std::endl;
                if (c.hasFormula()) {
                    if (c.formula().get().length() > 0)
                        cout << "cell " << ref.address() << " formula is |" << c.formula().get() << "|" << std::endl;
                    else    // NOTE: this state can not be created with OpenXLSX setter functions, but possibly by other libraries / programs
                        cout << "cell " << ref.address() << " has an empty formula" << std::endl;
                }
            }
        }
    }

    doc.save();
    doc.close();

    return 0;
}
