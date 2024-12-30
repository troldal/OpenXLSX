#include <cmath>
#include <iostream>
#include <string>

#include <OpenXLSX.hpp>
#include <external/pugixml/pugixml.hpp>

bool nowide_status()
{
#ifdef ENABLE_NOWIDE
    return true;
#else
    return false;
#endif
}

using namespace std;
using namespace OpenXLSX;

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #09: Cell assignment and access operations and some other features:\n";
    cout << " * XLCellReference::columnAsString, ::columnAsNumber, XLCoordinates::coordinatesFromAddress\n";
    cout << " * Worksheet name manipulations\n";
    cout << " * Range iteration using XLCellIterator::endReached() and showcasing overloaded operator<< for iostream\n";
    cout << " * XLCell and XLCellAssignable copy and move assignment operations\n";
    cout << " * XLCell and XLCellAssignable operator overloads\n";
    cout << " * Manipulate column width\n";
    cout << " * Iterate over XLWorksheet::rows()\n";
    cout << " * XLRow values assignment from/to vector, lists\n";
    cout << " * XLCellValue::typeAsString and ::getString\n";
    cout << " * Behavior of active worksheet when sheet setActive is used and when sheets are hidden\n";
    cout << " * XLRowDataIterator operators == and != (lhs and rhs being tested on both sides)\n";
    cout << "********************************************************************************\n";
    std::cout << "nowide is " << ( nowide_status() ? "enabled" : "disabled" ) << std::endl;

    // load example file
    XLDocument doc;
    doc.create("./Demo09.xlsx", XLForceOverwrite);

    std::string version = "12.345";
    doc.setProperty( XLProperty::AppVersion, version );
    std::cout << "doc.property(XLProperty::AppVersion) is \"" << doc.property(XLProperty::AppVersion) << "\"" << std::endl;

    cout << "\n=== Demonstrate XLCellReference::columnAsString, ::columnAsNumber, XLCoordinates::coordinatesFromAddress ===\n";
    XLCellReference ref("A1");
    std::cout << "column #761 translates to " << ref.columnAsString( 761 ) << std::endl;
    std::cout << "column ACG translates to " << ref.columnAsNumber( "ACG" ) << std::endl;

    uint32_t maxRows = OpenXLSX::MAX_ROWS;
    std::string address = "XFD" + std::to_string( maxRows ); // last allowed column and row
    XLCoordinates coords = XLCellReference::coordinatesFromAddress(address);
    std::cout << "coords from " << address << " become row " << coords.first << " and col " << coords.second << std::endl;

    // address = "XFE" + std::to_string( maxRows ); // last allowed column and row
    // coords = XLCellReference::coordinatesFromAddress(address);
    // std::cout << "coords from " << address << " become row " << coords.first << " and col " << coords.second << std::endl;

    // address = "XFD" + std::to_string( maxRows + 1 ); // last allowed column and row
    // coords = XLCellReference::coordinatesFromAddress(address);
    // std::cout << "coords from " << address << " become row " << coords.first << " and col " << coords.second << std::endl;

    // address = "12345"; // no column
    // coords = XLCellReference::coordinatesFromAddress(address);
    // std::cout << "coords from " << address << " become row " << coords.first << " and col " << coords.second << std::endl;

    // address = "AB"; // no row
    // coords = XLCellReference::coordinatesFromAddress(address);
    // std::cout << "coords from " << address << " become row " << coords.first << " and col " << coords.second << std::endl;

    address = "AA";
    uint16_t col = XLCellReference::columnAsNumber(address);
    std::cout << "columnAsNumber from " << address << " becomes col " << col << std::endl;

    // address = "XFE";
    // col = XLCellReference::columnAsNumber(address);
    // std::cout << "columnAsNumber from " << address << " becomes col " << col << std::endl;

    cout << "\n=== Worksheet name manipulations ===\n";
    auto const & wsNames = doc.workbook().worksheetNames();
    for( size_t idx = 0; idx < wsNames.size(); ++idx )
        cout << "worksheet name is " << wsNames[ idx ] << ", index is " << idx << endl;

    for (std::string name : doc.workbook().chartsheetNames()) // chart sheets are empty - just demonstrate how to access their names
        cout << "chartsheet name is \"" << name << "\"" << endl;

    auto wks = doc.workbook().worksheet( 1 );
    wks.setName( "Worksheet name test");
    std::cout << "worksheet name is " << wks.name() << std::endl;

    // assign some data to row 1 from a vector of strings:
    wks.row(1).values() = std::vector< std::string >{ "column A", "column B", "column C", "column D", "column E", "column F", "column G" };

    cout << "\n=== Range iteration using XLCellIterator::endReached() and showcasing overloaded operator<< for iostream ===\n";
    auto rng1 = wks.range(XLCellReference("A1"), XLCellReference("I2"));
    XLCellIterator i1 = rng1.begin();
    uint32_t count = 0;
    XLCellIterator i2 = i1;
    while( !i2.endReached() ) { ++i2; ++count; }
    std::cout << "i2.endReached() is " << ( i2.endReached() ? "true" : "false" ) << std::endl;

    i2 = i1;
    for( uint32_t i = 0; i < count * 0.66; ++i, ++i2 );
    std::cout << "distance between " << i1 << " and " << i2 << " is " << i1.distance(i2) << std::endl;

    while( !i2.endReached() ) ++i2;
    std::cout << "end iterator of rng1 is " << i2 << std::endl;


    cout << "\n=== XLCell and XLCellAssignable copy and move assignment operations ===\n";
    std::cout << "A1 value() is " << wks.cell("A1") << ", A1 formula() is " << wks.cell("A1").formula() << std::endl;
    std::cout << "B1 value() is " << wks.cell("B1") << ", B1 formula() is " << wks.cell("B1").formula() << std::endl;
    XLCell c = wks.cell("A1");   // XLCell move constructor
    std::cout << "c is " << c << std::endl;
    XLCell c2 = wks.cell("A1");  // XLCell move constructor
    std::cout << "c2 is " << c2 << std::endl;
    c2 = c;                      // XLCell copy assignment
    std::cout << "c2 is " << c2 << std::endl;
    try {
        // manual invocation of XLCell.copyFrom (automatically used by XLCellAssignable copy assignment operator)
        c2.copyFrom( c ); // copyFrom throws on a cell referencing the same XML cell node
        std::cout << "c2 is " << c2 << std::endl;
    }
    catch( XLInputError const & e ) {
        std::cout << "--ERROR: c2.copyFrom( c ) failed: " << e.what() << std::endl;
    }
    int iteration = 0;
    try {
        XLCell c3;
        while( iteration < 2 ) {
            ++iteration;
            c3.copyFrom( wks.cell("A1") ); // copyFrom throws on a cell referencing the same XML cell node
            std::cout << "c3 is " << c3 << std::endl;
        }
    }
    catch( XLInputError const & e ) {
        std::cout << "--ERROR: c3.copyFrom( c ) failed on iteration " << iteration << ": " << e.what() << std::endl;
    }

    // XLCellAssignable ca = wks.cell("A1"); // copy constructor
    XLCellAssignable ca;
    ca = c;                 // XLCellAssignable copy assignment -> refers to "A1"
    std::cout << "ca is " << ca << std::endl;

    ca = wks.cell("B1");    // XLCellAssignable move assignment -> assigns contents of "B1" to "A1"
    std::cout << "ca is " << ca << std::endl;
    XLCell x = ca;          // XLCell copy assignment
    std::cout << "x is " << x << std::endl;


    cout << "\n=== XLCell and XLCellAssignable operator overloads ===\n";
    wks.cell("A25" ).value() = "42";      // value assignment of any type
    wks.cell("A25" ).formula() = "=4+3";  // formula assignment from string
    std::cout << "A25 value() is " << wks.cell("A25") << ", A25 formula() is " << wks.cell("A25").formula() << std::endl;
    wks.cell("C11") = wks.cell("A25");    // copy assignment of all cell contents (like copy & paste in Excel)
    std::cout << "C11 value() is " << wks.cell("C11") << ", C11 formula() is " << wks.cell("C11").formula() << std::endl;

    std::cout << "D12    before cell copy: D12 value() is " << wks.cell("D12") << ", D12 formula() is " << wks.cell("D12").formula() << std::endl;
    wks.cell("D12") = wks.cell("A1");
    std::cout << "D12  after copy from A1: D12 value() is " << wks.cell("D12") << ", D12 formula() is " << wks.cell("D12").formula() << std::endl;
    wks.cell("D12") = wks.cell("C11");
    std::cout << "D12 after copy from C11: D12 value() is " << wks.cell("D12") << ", D12 formula() is " << wks.cell("D12").formula() << std::endl;

    // copy some cell contents around (value, attributes & formula):
    wks.cell("D13") = wks.cell("C11");
    wks.cell("D13") = wks.cell("A1");
    XLCellValue D13 = wks.cell("D13").value();
    std::cout << "D13 hasFormula() is " << ( wks.cell("D13").hasFormula() ? "true" : "false" ) << std::endl;
    std::cout << "D13 value() is " << D13.getString() << ", D13 formula() is " << wks.cell("D13").formula() << std::endl;

    cout << "\n=== Manipulate column width ===\n";
    std::cout << "wks.column( 3 ).width() is " << wks.column( 3 ).width() << std::endl;
    wks.column( 3 ).setWidth( 1 );
    std::cout << "wks.column( 3 ).width() is " << wks.column( 3 ).width() << std::endl;


    cout << "\n=== Iterate over XLWorksheet::rows() ===\n";
    // auto rng = wks.range(XLCellReference("A1"), XLCellReference("C7"));
    // auto it = rng.begin();
    // while( it != rng.end() ) {
    //     uint32_t row = it->cellReference().row();
    //     uint16_t col = it->cellReference().column();
    //     std::cout << "cell row is " << row << ", column is " << col << std::endl;
    //     ++it; // trigger XCellIterator::operator++
    // }
    std::cout << "wks.rowCount is " << wks.rowCount() << std::endl;
    std::cout << "wks.rows().rowCount is " << wks.rows().rowCount() << std::endl;
    auto row = wks.rows().begin();
    uint32_t limit = 55;
    while( row != wks.rows().end() && row->rowNumber() < limit ) { // iterate over worksheet rows with a safeguard limit
        std::cout << "row number is " << row->rowNumber() << std::endl;
        ++row;
    }

    cout << "\n=== Assignment of row values from/to vector, lists ===\n";
    // wks.row(2).values() = std::vector<int> { 1, 2, 3, 4, 5, 6, 7, 8 };
    wks.row(2).values() = std::vector<bool> { true, true, false, false, true, false, true };

    std::list< XLCellValue > l2;
    for( int i = 0; i < 10; ++i ) {
        l2.push_back( std::move( XLCellValue(i) ) );
    }
    wks.row(3).values() = l2; // assign a std::list of values to a worksheet row
    // wks.row(3).values() = wks.row(2).cells();

    std::vector< XLCellValue > vec1 = wks.row(1).values(); // get row values as a std::vector of XLCellValue
    for( auto v : vec1 )
        std::cout << "vec1 value is " << v.getString() << std::endl;

    std::vector< XLCellValue > vec2;
    vec2 = wks.row(4).values();
    std::cout << "vec2.size() before assigning row(4).values() = vec1 is " << vec2.size() << std::endl;
    for( auto v : vec2 )
        std::cout << "vec2 value before assigning row(4).values() = vec1 is " << v.getString() << std::endl;

    wks.row(4).values() = vec1; // assign a std::vector of XLCellValue to row values

    vec2 = wks.row(4).values();
    std::cout << "vec2.size() AFTER assigning row(4).values() = vec1 is " << vec2.size() << std::endl;
    for( auto v : vec2 )
        std::cout << "vec2 value AFTER assigning row(4).values() = vec1 is " << v.getString() << std::endl;


    cout << "\n=== Testing for cell data types & getting a string for any type ===\n";
    XLCellValue B3 = wks.cell("B3").value();
    cout << "Cell B3 (" << B3.typeAsString() << "): " << B3.getString() << endl;

    XLCellValue G2 = wks.cell("G2").value();
    cout << "Cell G2 (" << G2.typeAsString() << "): " << G2.getString() << endl;

    XLCellValue G1 = wks.cell("G1").value();
    std::string vG1 = "";
    try {
        vG1 = G1.get< std::string >();
    }
    catch( ... ) {
        std::cout << "failed to G1.get< std::string >(), attempting the new std::visit getString..." << std::endl;
        vG1 = G1.getString();
    }
    std::string sG1 = G1.typeAsString();

    cout << "Cell G1 (" << sG1 << "): " << vG1 << endl;

    XLCellValue A1 = wks.cell("A1").value();
    cout << "Cell A1 (" << A1.typeAsString() << "): " << A1.getString() << endl;
    XLCellValue B1 = wks.cell("B1").value();
    cout << "Cell B1 (" << B1.typeAsString() << "): " << B1.getString() << endl;
    XLCellValue C1 = wks.cell("C1").value();
    cout << "Cell C1 (" << C1.typeAsString() << "): " << C1.getString() << endl;
    // if, due to linting (whitespaces), access to cell fails, an empty cell node gets added to the worksheet, returning the empty value
    //  --> the correct value is still there, but somehow has the wrong node reference
    XLCellValue D1 = wks.cell("D1").value();
    cout << "Cell D1 (" << D1.typeAsString() << "): " << D1.getString() << endl;
    XLCellValue E1 = wks.cell("E1").value();
    cout << "Cell E1 (" << E1.typeAsString() << "): " << E1.getString() << endl;
    XLCellValue F1 = wks.cell("F1").value();
    cout << "Cell F1 (" << F1.typeAsString() << "): " << F1.getString() << endl;


    // std::list< XLCellValue > l1 = wks.row(1).values();
    // for( auto v : l1 )
    //     std::cout << "l1 value is " << v.getString() << std::endl;

    XLCellValue Z35 = wks.cell("Z35").value();
    cout << "Cell Z35 (" << Z35.typeAsString() << "): " << Z35.getString() << endl;
    A1 = wks.cell("A35").value();


    cout << "\n=== Demonstrate behavior of active worksheet when sheet setActive is used and when sheets are hidden ===\n";
    std::cout << "doc.workbook().sheetCount() is " << doc.workbook().sheetCount() << std::endl;
    for( size_t i = 1; i <= doc.workbook().sheetCount(); ++i )
        std::cout << "workbook.worksheet(" << i << ").name() is " << doc.workbook().worksheet(i).name() << std::endl;

    std::string newSheetName;
    newSheetName = "Sheet2"; if( !doc.workbook().sheetExists( newSheetName ) ) doc.workbook().addWorksheet( newSheetName );
    auto wks2 = doc.workbook().worksheet( newSheetName );
    newSheetName = "Sheet3"; if( !doc.workbook().sheetExists( newSheetName ) ) doc.workbook().addWorksheet( newSheetName );
    auto wks3 = doc.workbook().worksheet( newSheetName );

    std::cout << "wks.isActive() is " << ( wks.isActive() ? "true" : "false" ) << std::endl;
    std::cout << "wks2.isActive() is " << ( wks2.isActive() ? "true" : "false" ) << std::endl;
    std::cout << "wks3.isActive() is " << ( wks3.isActive() ? "true" : "false" ) << std::endl;

    for( int sheetIndex = 1; sheetIndex <= 3; ++sheetIndex ) {
        doc.workbook().sheet( sheetIndex ).setSelected(false);
        doc.workbook().sheet( sheetIndex ).setVisibility(XLSheetState::Visible);
    }

    wks.setActive();
    wks2.setActive();
    wks3.setActive();
    std::cout << "   wks.isActive() is " << ( wks.isActive() ? "true" : "false" ) << std::endl;
    std::cout << "   wks2.isActive() is " << ( wks2.isActive() ? "true" : "false" ) << std::endl;
    std::cout << "   wks3.isActive() is " << ( wks3.isActive() ? "true" : "false" ) << std::endl;
    // XLSheetState::Visible, XLSheetState::Hidden, XLSheetState::VeryHidden
    // wks.setVisibility( XLSheetState::Visible );
    std::cout << "Setting wks2 to active..." << ( wks2.setActive() ? " succeeded." : "failed!" ) << std::endl;
    std::cout << "   wks.isActive() is " << ( wks.isActive() ? "true" : "false" ) << std::endl;
    std::cout << "   wks2.isActive() is " << ( wks2.isActive() ? "true" : "false" ) << std::endl;
    std::cout << "   wks3.isActive() is " << ( wks3.isActive() ? "true" : "false" ) << std::endl;

    std::cout << "Hiding wks2" << std::endl;
    wks2.setVisibility( XLSheetState::VeryHidden );
    std::cout << "   wks.isActive() is " << ( wks.isActive() ? "true" : "false" ) << std::endl;
    std::cout << "   wks2.isActive() is " << ( wks2.isActive() ? "true" : "false" ) << std::endl;
    std::cout << "   wks3.isActive() is " << ( wks3.isActive() ? "true" : "false" ) << std::endl;

    std::cout << "Setting wks3 to active... " << ( wks3.setActive() ? "succeeded." : "failed!" ) << std::endl;
    std::cout << "   wks.isActive() is " << ( wks.isActive() ? "true" : "false" ) << std::endl;
    std::cout << "   wks2.isActive() is " << ( wks2.isActive() ? "true" : "false" ) << std::endl;
    std::cout << "   wks3.isActive() is " << ( wks3.isActive() ? "true" : "false" ) << std::endl;

    std::cout << "Setting wks2 to active... " << ( wks2.setActive() ? "succeeded." : "failed!" ) << std::endl;
    std::cout << "   wks.isActive() is " << ( wks.isActive() ? "true" : "false" ) << std::endl;
    std::cout << "   wks2.isActive() is " << ( wks2.isActive() ? "true" : "false" ) << std::endl;
    std::cout << "   wks3.isActive() is " << ( wks3.isActive() ? "true" : "false" ) << std::endl;

    doc.workbook().setSheetIndex( "Sheet3", 2 );
    // How to get a sheet type from the shared sheet index (indexes both worksheets and chartsheets)
    for( int sheetIndex = 1; sheetIndex <= 3; ++sheetIndex ) {
        auto sheetType = doc.workbook().typeOfSheet( sheetIndex );
        std::string sheetTypeString = "(unknown)";
        if( sheetType == XLSheetType::Worksheet ) sheetTypeString = "Worksheet";
        else if( sheetType == XLSheetType::Chartsheet ) sheetTypeString = "Chartsheet";
        std::cout << "sheet( " << sheetIndex << " ) of type " + sheetTypeString + " has name " << doc.workbook().sheet( sheetIndex ).name()
            << " and is" << ( doc.workbook().sheet( sheetIndex ).isSelected() ? "" : " NOT" ) << " selected" << std::endl;
        // doc.workbook().sheet( sheetIndex ).print( std::cout );
    }

    // output status before saving workbook:
    std::cout << "sheetNames is:" << std::endl;
    for (std::string name : doc.workbook().sheetNames() )
        std::cout << "  " << name << std::endl;

    std::cout << "worksheetNames is:" << std::endl;
    for (std::string name : doc.workbook().worksheetNames() )
        std::cout << "  " << name << std::endl;

    std::cout << "chartsheetNames is:" << std::endl;
    for (std::string name : doc.workbook().chartsheetNames() )
        std::cout << "  " << name << std::endl;

    cout << "\n=== Demonstrate operators == and != for XLRowDataIterator (lhs and rhs being tested on both sides) ===\n";
    // uint32_t rowNo = 1'048'576; // this would result in wks.row( rowNo ).cells().begin() returning an end iterator
    uint32_t rowNo = 1;
    std::cout << "wks.row( " << rowNo << " ).cells().size() is " << wks.row( rowNo ).cells().size() << std::endl;
    XLRowDataIterator r1 = wks.row( rowNo ).cells().begin();
    XLRowDataIterator rEnd = wks.row( rowNo ).cells().end();
    std::cout << "r1 == rEnd evaluates to " << ( r1 == rEnd ? "true" : "false" ) << std::endl;
    std::cout << "rEnd == r1 evaluates to " << ( rEnd == r1 ? "true" : "false" ) << std::endl;
    std::cout << "r1 == r1 evaluates to " << ( r1 == r1 ? "true" : "false" ) << std::endl;
    std::cout << "rEnd == rEnd evaluates to " << ( rEnd == rEnd ? "true" : "false" ) << std::endl;
    std::cout << std::endl;
    std::cout << "r1 != rEnd evaluates to " << ( r1 != rEnd ? "true" : "false" ) << std::endl;
    std::cout << "rEnd != r1 evaluates to " << ( rEnd != r1 ? "true" : "false" ) << std::endl;
    std::cout << "r1 != r1 evaluates to " << ( r1 != r1 ? "true" : "false" ) << std::endl;
    std::cout << "rEnd != rEnd evaluates to " << ( rEnd != rEnd ? "true" : "false" ) << std::endl;

    doc.save();    // implies XLForceOverwrite
    doc.close();

    return 0;
}
