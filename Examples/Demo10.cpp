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

int main( int argc, char *argv[] )
{
	std::cout << "nowide is " << ( nowide_status() ? "enabled" : "disabled" ) << std::endl;

	/*
	 * TODO: align XLStyles.cpp with known documentation of xl/styles.xml as far as support is desired
	 *  * XLFills: support patternFill patternType (XLFillPattern)
	 *  * XLFonts: support scheme major/minor/none (<font><scheme>major</scheme></font>)
	 *  * XLStyles ::create functions: implement good default style properties for all styles
	 *  * support "colors"
	 *  * support "dxfs"
	 *  * support "tableStyles"
	 * TODO TBD: permit setting a format reference for shared strings
	 * 
	 * DONE: XLFills: support gradientFill
	 * DONE: support complete for: "borders", "cellStyles", "cellStyleXfs", "cellXfs", "numFmts"
	 * DONE: support nearly complete (exceptions above) for: "fonts"
	 * DONE: support for xfId / setXfId in XLStyles::cellFormats, but not(!) in XLStyles::cellStyleFormats
	 * DONE: format support for XLCells (cellFormat / setCellFormat)
	 * DONE: format support for XLRows: <row> attributes s (same as cell attribute) and customFormat (=true/false)
	 *       XLWorksheet::setRowFormat overwrites existing cell formats, and
	 *         TODO TBD row format is then used by OpenXLSX for new cells created in that row
	 * DONE: format support for XLColumns: <col> attributes style (same as cell attribute s)
	 *       XLWorksheet::setColumnFormat overwrites existing cell formats, and
	 *         TODO TBD column format is then used by OpenXLSX as default for new cells created in that column when they do not have a row style
	 * DONE: update style elements "count" attribute for all arrays when saving
	 */
	
	// create new workbook
	XLDocument doc{};
	doc.create("./Demo10.xlsx", XLForceOverwrite);
	// doc.open("./Demo10.xlsx");

	std::cout << "doc.name() is " << doc.name() << std::endl;


	XLNumberFormats & numberFormats = doc.styles().numberFormats();
	XLFonts & fonts = doc.styles().fonts();
	XLFills & fills = doc.styles().fills();
	XLBorders & borders = doc.styles().borders();
	XLCellFormats & cellStyleFormats = doc.styles().cellStyleFormats();
	XLCellFormats & cellFormats = doc.styles().cellFormats();
	XLCellStyles & cellStyles = doc.styles().cellStyles();

	XLWorksheet wks = doc.workbook().worksheet(1);
	XLStyleIndex cellFormatIndexA2 = wks.cell("A2").cellFormat();                 // get index of cell format
	XLStyleIndex fontIndexA2 = cellFormats[ cellFormatIndexA2 ].fontIndex();      // get index of used font

	XLStyleIndex newFontIndex = fonts.create( fonts[ fontIndexA2 ] );             // create a new font based on used font
	XLStyleIndex newCellFormatIndex                                               // create a new cell format
		= cellFormats.create( cellFormats[ cellFormatIndexA2 ] );                  //              based on used format

	cellFormats[ newCellFormatIndex ].setFontIndex( newFontIndex );         // assign the new font to the new cell format
	wks.cell("A2").setCellFormat( newCellFormatIndex );                     // assign the new cell format to the cell

	//                                                                      // change the new font style:
	fonts[ newFontIndex ].setFontName("Arial");
	fonts[ newFontIndex ].setFontSize(14);
	XLColor red   ( "ffff0000" );
	XLColor green ( "ff00ff00" );
	XLColor blue  ( "ff0000ff" );
	XLColor yellow( "ffffff00" );
	fonts[ newFontIndex ].setFontColor( green );
	fonts[ newFontIndex ].setBold();                                        // bold
	fonts[ newFontIndex ].setItalic();                                      // italic
	fonts[ newFontIndex ].setStrikethrough();                               // strikethrough
	fonts[ newFontIndex ].setUnderline(XLUnderlineDouble);                  // underline: XLUnderlineSingle, XLUnderlineDouble, XLUnderlineNone
	fonts[ newFontIndex ].setUnderline(XLUnderlineNone);                    // TEST: un-set the underline status

	cellFormats[ newCellFormatIndex ].setApplyFont( true );                 // apply the font (seems to do nothing)

	wks.cell("A25" ) = "42"; // NEW functionality: XLCellAssignable now has a direct value assignment overload
	std::cout << "wks A25 value is " << wks.cell("A25" ) << std::endl;

	XLCellRange cellRange = wks.range("A3", "Z10"); // NEW functionality: range based on cell reference strings
	std::cout << "range " << cellRange.address()                      // NEW: cell range address
		<< " with top left " << cellRange.topLeft().address()          // NEW: cell range topLeft cell reference
		<< " and bottom right " << cellRange.bottomRight().address()   // NEW: cell range bottomRight cell reference
		<< " contains " << cellRange.numRows() << " rows, " << cellRange.numColumns() << " columns" << std::endl;

	cellRange = wks.range("A1:XFD100"); // NEW functionality: range based on range reference
	std::cout << "range " << cellRange.address() << " contains " << cellRange.numRows() << " rows, " << cellRange.numColumns() << " columns" << std::endl;

	XLStyleIndex fontBold = fonts.create();
	fonts[ fontBold ].setBold();
	XLStyleIndex boldDefault = cellFormats.create();
	cellFormats[ boldDefault ].setFontIndex( fontBold );

	cellRange = wks.range("A5:Z5");
	cellRange = "this is a bold range"; // NEW functionality: XLCellRange now has a direct value assignment overload
	cellRange.setFormat( boldDefault ); // NEW functionality: XLCellRange supports setting format for all cells in the range

	XLStyleIndex fontItalic = fonts.create();
	fonts[ fontItalic ].setItalic();
	XLStyleIndex italicDefault = cellFormats.create();
	cellFormats[ italicDefault ].setFontIndex( fontItalic );

	cellRange = wks.range("E1:E100");
	cellRange = "this is an italic range";
	cellRange.setFormat( italicDefault );

	cellRange = wks.range("A10:Z11");
	cellRange = "lalala"; // create cells if they do not exist by assigning values to them
	XLRow row = wks.row(10);
	row.cells() = "this is an italic row"; // NEW functionality: overwrite value for all *existing* cells in row
	wks.setRowFormat( row.rowNumber(), italicDefault );  // NEW functionality: set format for all (existing) cells in row, and row format to be used as default for future cells

	cellRange = wks.range("J1:J100");
	cellRange = "this is a bold column";
	XLColumn col = wks.column("J");      // NEW functionality: XLWorksheet::column now supports getting a column from a column reference
	wks.setColumnFormat( XLCellReference::columnAsNumber("J"), boldDefault ); // NEW functionality: set format for all (existing) cells in column, and column format to be used as default for future cells

	wks.setRowFormat( 11, newCellFormatIndex );  // set row 11 to bold italic underline strikethrough font with color as formatted above
	XLStyleIndex redFormat = cellFormats.create( cellFormats[ newCellFormatIndex ] );
	XLStyleIndex redFont = fonts.create();
	fonts[ redFont ].setFontColor( red );
	cellFormats[ redFormat ].setFontIndex( redFont );
	cellRange = wks.range("B2:P5");
	cellRange = "red range";
	wks.cell(cellRange.topLeft()).value() = "red range " + cellRange.address() + " top left cell";
	wks.cell(cellRange.bottomRight()).value() = "red range " + cellRange.address() + " bottom right cell";
	cellRange.setFormat( redFormat );

	// // get the relevant style objects for easy access
	// XLCellFormats & cellFormats = doc.styles().cellFormats();
	// XLFonts & fonts = doc.styles().fonts();
	// XLFills & fills = doc.styles().fills();

	XLCellRange myCellRange = wks.range("B20:P25");           // create a new range for formatting
	myCellRange = "TEST COLORS";                              // write some values to the cells so we can see format changes
	XLStyleIndex baseFormat = wks.cell("B20").cellFormat();   // determine the style used in B20
	XLStyleIndex newCellStyle = cellFormats.create( cellFormats[ baseFormat ] ); // create a new style based on the style in B20
	XLStyleIndex newFontStyle = fonts.create(fonts[ cellFormats[ baseFormat ].fontIndex() ]); // create a new font style based on the used font
	XLStyleIndex newFillStyle = fills.create(fills[ cellFormats[ baseFormat ].fillIndex() ]); // create a new fill style based on the used fill

	fonts[ newFontStyle ].setFontColor( green );
	cellFormats[ newCellStyle ].setFontIndex( newFontStyle );

	fills[ newFillStyle ].setPatternType    ( XLPatternSolid );           // "solid"
	fills[ newFillStyle ].setColor          ( yellow );
	// fills[ newFillStyle ].setBackgroundColor( XLColor( "ff0000ff" ) );    // blue, setBackgroundColor only makes sense with gradient fills
	cellFormats[ newCellStyle ].setFillIndex( newFillStyle );

	myCellRange.setFormat( newCellStyle ); // assign the new format to the full range of cells

	// enable testBasics = true to create/modify at least one entry in each known styles array
	// disable testBasics = false to stop the demo execution here and save the document
	bool testBasics = true;

	if( testBasics ) {
		std::cout << "   numberFormats.count() is " <<    numberFormats.count() << std::endl;
		std::cout << "           fonts.count() is " <<            fonts.count() << std::endl;
		std::cout << "           fills.count() is " <<            fills.count() << std::endl;
		std::cout << "         borders.count() is " <<          borders.count() << std::endl;
		std::cout << "cellStyleFormats.count() is " << cellStyleFormats.count() << std::endl;
		std::cout << "     cellFormats.count() is " <<      cellFormats.count() << std::endl;
		std::cout << "      cellStyles.count() is " <<       cellStyles.count() << std::endl;


		// enable createAll = true to create a new entry in each known styles array and modify that
		// disable createAll = false to modify the last entry in each styles array instead
		bool createAll = true;

		XLStyleIndex nodeIndex;

		nodeIndex = ( createAll ? numberFormats.create() : numberFormats.count() - 1 );
		numberFormats[ nodeIndex ].setNumberFormatId( 15 );
		numberFormats[ nodeIndex ].setFormatCode( "fifteen" );
		std::cout << "numberFormats[ " << nodeIndex << " ] summary: " << numberFormats[ nodeIndex ].summary() << std::endl;

		nodeIndex = ( createAll ? fonts.create() : fonts.count() - 1 );
		fonts[ nodeIndex ].setFontName( "Twelve" );
		fonts[ nodeIndex ].setFontCharset( 13 );
		fonts[ nodeIndex ].setFontFamily( 14 );
		fonts[ nodeIndex ].setFontSize( 15 );
		fonts[ nodeIndex ].setFontColor( XLColor( XLDefaultFontColor ) );
		fonts[ nodeIndex ].setBold();                                        // bold
		fonts[ nodeIndex ].setItalic();                                      // italic
		fonts[ nodeIndex ].setStrikethrough();                               // strikethrough
		fonts[ nodeIndex ].setUnderline(XLUnderlineDouble);                  // underline: XLUnderlineSingle, XLUnderlineDouble, XLUnderlineNone
		fonts[ nodeIndex ].setOutline();
		fonts[ nodeIndex ].setShadow();
		fonts[ nodeIndex ].setCondense();
		fonts[ nodeIndex ].setExtend();
		std::cout << "fonts[ " << nodeIndex << " ] summary: " << fonts[ nodeIndex ].summary() << std::endl;

		// XLFill -> XLPatternFill test:
		nodeIndex = ( createAll ? fills.create() : fills.count() - 1 );
		fills[ nodeIndex ].setFillType( XLPatternFill );
		fills[ nodeIndex ].setPatternType( XLDefaultPatternType );
		fills[ nodeIndex ].setColor          ( XLColor( XLDefaultPatternFgColor ) );
		fills[ nodeIndex ].setBackgroundColor( XLColor( XLDefaultPatternBgColor ) );
		std::cout << "fills[ " << nodeIndex << " ] summary: " << fills[ nodeIndex ].summary() << std::endl;

		// XLFill -> XLGradientFill test:
		nodeIndex = ( createAll ? fills.create() : fills.count() - 1 );
		fills[ nodeIndex ].setFillType( XLPatternFill );
		fills[ nodeIndex ].setFillType( XLGradientFill, XLForceFillType );
		fills[ nodeIndex ].setGradientType( XLGradientLinear );
		XLGradientStops stops = fills[ nodeIndex ].stops();

		// first XLGradientStop
		XLStyleIndex stopIndex = stops.create();
		XLDataBarColor stopColor = stops[ stopIndex ].color();
		stops[ stopIndex ].setPosition(3.14);
		stopColor.setRgb( green );
		stopColor.setTint( 1.01 );
		stopColor.setAutomatic();
		stopColor.setIndexed( 77 );
		stopColor.setTheme( 79 );

		// second XLGradientStop
		stopIndex = stops.create(stops[stopIndex]); // create another stop using previous stop as template
		stopColor.set( yellow );
		std::cout << "fills[ " << nodeIndex << " ] summary: " << fills[ nodeIndex ].summary() << std::endl;

		nodeIndex = ( createAll ? borders.create() : borders.count() - 1 );
		borders[ nodeIndex ].setDiagonalUp( false );
		borders[ nodeIndex ].setDiagonalDown( true );
		borders[ nodeIndex ].setOutline( true );
		borders[ nodeIndex ].setLeft      ( XLLineStyleThin,   XLColor( "ff112222" ) );
		borders[ nodeIndex ].setRight     ( XLLineStyleMedium, XLColor( "ff113333" ) );
		borders[ nodeIndex ].setTop       ( XLLineStyleDashed, XLColor( "ff114444" ) );
		borders[ nodeIndex ].setBottom    ( XLLineStyleDotted, XLColor( "ff222222" ), 0.25 );
		borders[ nodeIndex ].setDiagonal  ( XLLineStyleThick,  XLColor( "ff332222" ), -0.25 );
		borders[ nodeIndex ].setVertical  ( XLLineStyleDouble, XLColor( "ff443322" ) );
		borders[ nodeIndex ].setHorizontal( XLLineStyleHair,   XLColor( "ff446688" ) );
		std::cout << "#1 borders[ " << nodeIndex << " ] summary: " << borders[ nodeIndex ].summary() << std::endl;

		// test additional line styles with setter function:
		borders[ nodeIndex ].setLine      ( XLLineLeft,     XLLineStyleMediumDashed,     XLColor( XLDefaultFontColor ) );
		borders[ nodeIndex ].setLine      ( XLLineRight,    XLLineStyleDashDot,          XLColor( XLDefaultFontColor ), 1.0 );
		borders[ nodeIndex ].setLine      ( XLLineTop,      XLLineStyleMediumDashDot,    XLColor( XLDefaultFontColor ) );
		borders[ nodeIndex ].setLine      ( XLLineBottom,   XLLineStyleDashDotDot,       XLColor( XLDefaultFontColor ), -0.1 );
		borders[ nodeIndex ].setLine      ( XLLineDiagonal, XLLineStyleMediumDashDotDot, XLColor( XLDefaultFontColor ), -1.0 );
		borders[ nodeIndex ].setLine      ( XLLineVertical, XLLineStyleSlantDashDot,     XLColor( XLDefaultFontColor ) );
		std::cout << "#2 borders[ " << nodeIndex << " ] summary: " << borders[ nodeIndex ].summary() << std::endl;

		// test direct access to line properties (read: the XLDataBarColor)
		XLLine leftLine = borders[ nodeIndex ].left();
		XLDataBarColor leftLineColor = leftLine.color();
		leftLineColor.setRgb( blue );
		leftLineColor.setTint( -0.2 );
		leftLineColor.setAutomatic();
		leftLineColor.setIndexed( 606 );
		leftLineColor.setTheme( 707 );
		std::cout << "#2 borders, leftLine summary: " << leftLine.summary() << std::endl;
		// unset some properties
		leftLineColor.setTint( 0 );
		leftLineColor.setAutomatic( false );
		leftLineColor.setIndexed( 0 );
		leftLineColor.setTheme( 0 );
		std::cout << "#2 borders, leftLineColor summary: " << leftLineColor.summary() << std::endl;

		nodeIndex = ( createAll ? cellStyleFormats.create() : cellStyleFormats.count() - 1 );
		cellStyleFormats[ nodeIndex ].setNumberFormatId( 5 );
		cellStyleFormats[ nodeIndex ].setFontIndex( 6 );
		cellStyleFormats[ nodeIndex ].setFillIndex( 7 );
		cellStyleFormats[ nodeIndex ].setBorderIndex( 8 );
		cellStyleFormats[ nodeIndex ].setApplyFont(true);
		cellStyleFormats[ nodeIndex ].setApplyFill(false);
		cellStyleFormats[ nodeIndex ].setApplyBorder(false);
		cellStyleFormats[ nodeIndex ].setApplyAlignment(true);
		cellStyleFormats[ nodeIndex ].setApplyProtection(false);
		cellStyleFormats[ nodeIndex ].setLocked(true);
		cellStyleFormats[ nodeIndex ].setHidden(false);
		cellStyleFormats[ nodeIndex ].alignment(XLCreateIfMissing).setHorizontal(XLAlignRight);
		cellStyleFormats[ nodeIndex ].alignment(XLCreateIfMissing).setVertical(XLAlignCenter);
		cellStyleFormats[ nodeIndex ].alignment(XLCreateIfMissing).setTextRotation(255);
		cellStyleFormats[ nodeIndex ].alignment(XLCreateIfMissing).setWrapText(true);
		cellStyleFormats[ nodeIndex ].alignment(XLCreateIfMissing).setIndent(256);
		cellStyleFormats[ nodeIndex ].alignment(XLCreateIfMissing).setRelativeIndent(-255);
		cellStyleFormats[ nodeIndex ].alignment(XLCreateIfMissing).setJustifyLastLine(true);
		cellStyleFormats[ nodeIndex ].alignment(XLCreateIfMissing).setShrinkToFit(false);
		cellStyleFormats[ nodeIndex ].alignment(XLCreateIfMissing).setReadingOrder(XLReadingOrderLeftToRight);

		try { cellStyleFormats[ nodeIndex ].setXfId(9); } // TEST: this should throw
		catch (XLException const & e) {
			std::cout << "cellStyleFormats NOMINAL EXCEPTION: " << e.what() << std::endl;
		}
		try { (void) cellStyleFormats[ nodeIndex ].xfId(); } // TEST: this should throw
		catch (XLException const & e) {
			std::cout << "cellStyleFormats NOMINAL EXCEPTION: " << e.what() << std::endl;
		}

		std::cout << "cellStyleFormats[ " << nodeIndex << " ] summary: " << cellStyleFormats[ nodeIndex ].summary() << std::endl;

		nodeIndex = ( createAll ? cellFormats.create() : cellFormats.count() - 1 );
		cellFormats[ nodeIndex ].setNumberFormatId( 22 );
		cellFormats[ nodeIndex ].setFontIndex( 23 );
		cellFormats[ nodeIndex ].setFillIndex( 24 );
		cellFormats[ nodeIndex ].setBorderIndex( 25 );
		cellFormats[ nodeIndex ].setApplyNumberFormat(true);
		cellFormats[ nodeIndex ].setApplyFont(false);
		cellFormats[ nodeIndex ].setApplyFill(true);
		cellFormats[ nodeIndex ].setApplyBorder(true);
		cellFormats[ nodeIndex ].setApplyAlignment(false);
		cellFormats[ nodeIndex ].setApplyProtection(true);
		cellFormats[ nodeIndex ].setQuotePrefix(true);
		cellFormats[ nodeIndex ].setPivotButton(true);
		cellFormats[ nodeIndex ].setLocked(false);
		cellFormats[ nodeIndex ].setHidden(true);
		cellFormats[ nodeIndex ].alignment(XLCreateIfMissing).setHorizontal(XLAlignGeneral);
		cellFormats[ nodeIndex ].alignment(XLCreateIfMissing).setVertical(XLAlignTop);
		cellFormats[ nodeIndex ].alignment(XLCreateIfMissing).setTextRotation(370);
		cellFormats[ nodeIndex ].alignment(XLCreateIfMissing).setWrapText(false);
		cellFormats[ nodeIndex ].alignment(XLCreateIfMissing).setIndent(37);
		cellFormats[ nodeIndex ].alignment(XLCreateIfMissing).setShrinkToFit(true);
		cellFormats[ nodeIndex ].setXfId(26);
		std::cout << "cellFormats[ " << nodeIndex << " ] summary: " << cellFormats[ nodeIndex ].summary() << std::endl;

		nodeIndex = ( createAll ? cellStyles.create() : cellStyles.count() - 1 );
		cellStyles[ nodeIndex ].setName( "test cellStyle" );
		cellStyles[ nodeIndex ].setXfId( 77 );
		cellStyles[ nodeIndex ].setBuiltinId( 7 );
		cellStyles[ nodeIndex ].setOutlineStyle( 8 );
		cellStyles[ nodeIndex ].setHidden( true );
		cellStyles[ nodeIndex ].setCustomBuiltin( true );
		std::cout << "cellStyles[ " << nodeIndex << " ] summary: " << cellStyles[ nodeIndex ].summary() << std::endl;


		std::cout << "   numberFormats.count() is " <<    numberFormats.count() << std::endl;
		std::cout << "           fonts.count() is " <<            fonts.count() << std::endl;
		std::cout << "           fills.count() is " <<            fills.count() << std::endl;
		std::cout << "         borders.count() is " <<          borders.count() << std::endl;
		std::cout << "cellStyleFormats.count() is " << cellStyleFormats.count() << std::endl;
		std::cout << "     cellFormats.count() is " <<      cellFormats.count() << std::endl;
		std::cout << "      cellStyles.count() is " <<       cellStyles.count() << std::endl;
	} // end if( testBasics )

	
	doc.saveAs("./Demo10.xlsx", XLForceOverwrite);
	doc.close();

	return 0;
}
