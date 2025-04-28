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
	cout << "DEMO PROGRAM #10: XLStyles Usage and XLSheet conditional formatting\n";
	cout << "********************************************************************************\n";

	std::cout << "nowide is " << ( nowide_status() ? "enabled" : "disabled" ) << std::endl;

	/*
	 * TODO: align XLStyles.cpp with known documentation of xl/styles.xml as far as support is desired
	 *  * XLAlignmentStyle: check / throw if vertical alignments are used as horizontal and vice versa
	 *  * XLStyles ::create functions: implement good default style properties for all styles
	 *  * support "colors"
	 *  * support "dxfs"
	 *  * support "tableStyles"
	 * TODO TBD: permit setting a format reference for shared strings
	 * 
	 * DONE: XLFills: support patternFill patternType (XLFillPattern)
	 * DONE: XLFonts: support scheme major/minor/none (<font><scheme val="major"/></font>)
	 * DONE: XLFonts: support vertical align run style (<font><vertAlign val="subscript"/></font>)
	 * DONE: XLFills: support gradientFill
	 * DONE: support complete for: "borders", "cellStyles", "cellStyleXfs", "cellXfs", "numFmts"
	 * DONE: support nearly complete (exceptions above) for: "fonts"
	 * DONE: support for xfId / setXfId in XLStyles::cellFormats, but not(!) in XLStyles::cellStyleFormats
	 * DONE: format support for XLCells (cellFormat / setCellFormat)
	 * DONE: format support for XLRows: <row> attributes s (same as cell attribute) and customFormat (=true/false)
	 *       XLWorksheet::setRowFormat overwrites existing cell formats, and
	 *         row format is then used by OpenXLSX for new cells created in that row
	 * DONE: format support for XLColumns: <col> attributes style (same as cell attribute s)
	 *       XLWorksheet::setColumnFormat overwrites existing cell formats, and
	 *         column format is then used by OpenXLSX as default for new cells created in that column when they do not have a row style
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

	fills[ newFillStyle ].setPatternType    ( XLPatternNone );
	fills[ newFillStyle ].setPatternType    ( XLPatternSolid );
	// fills[ newFillStyle ].setPatternType    ( XLPatternMediumGray );
	// fills[ newFillStyle ].setPatternType    ( XLPatternDarkGray );
	// fills[ newFillStyle ].setPatternType    ( XLPatternLightGray );
	// fills[ newFillStyle ].setPatternType    ( XLPatternDarkHorizontal );
	// fills[ newFillStyle ].setPatternType    ( XLPatternDarkVertical );
	// fills[ newFillStyle ].setPatternType    ( XLPatternDarkDown );
	// fills[ newFillStyle ].setPatternType    ( XLPatternDarkUp );
	// fills[ newFillStyle ].setPatternType    ( XLPatternDarkGrid );
	// fills[ newFillStyle ].setPatternType    ( XLPatternDarkTrellis );
	// fills[ newFillStyle ].setPatternType    ( XLPatternLightHorizontal );
	// fills[ newFillStyle ].setPatternType    ( XLPatternLightVertical );
	// fills[ newFillStyle ].setPatternType    ( XLPatternLightDown );
	// fills[ newFillStyle ].setPatternType    ( XLPatternLightUp );
	// fills[ newFillStyle ].setPatternType    ( XLPatternLightGrid );
	// fills[ newFillStyle ].setPatternType    ( XLPatternLightTrellis );
	// fills[ newFillStyle ].setPatternType    ( XLPatternGray125 );
	// fills[ newFillStyle ].setPatternType    ( XLPatternGray0625 );
	fills[ newFillStyle ].setColor          ( yellow );
	fills[ newFillStyle ].setBackgroundColor( blue );    // blue, setBackgroundColor only makes sense with gradient fills
	cellFormats[ newCellStyle ].setFillIndex( newFillStyle );

	myCellRange.setFormat( newCellStyle ); // assign the new format to the full range of cells


	// XLFill -> XLGradientFill test:
	myCellRange = wks.range("B30:P35");           // create a new range for formatting
	myCellRange = "TEST GRADIENT";                            // write some values to the cells so we can see format changes
	baseFormat = wks.cell("B30").cellFormat();   // determine the style used in B20
	newCellStyle = cellFormats.create( cellFormats[ baseFormat ] ); // create a new style based on the style in B20
	XLStyleIndex testGradientIndex = fills.create(fills[ cellFormats[ baseFormat ].fillIndex() ]); // create a new fill style based on the used fill

	fills[ testGradientIndex ].setFillType( XLPatternFill );                   // set the "wrong" fill type
	fills[ testGradientIndex ].setFillType( XLGradientFill, XLForceFillType ); // override that with a gradientFill
	fills[ testGradientIndex ].setGradientType( XLGradientLinear );            // the gradient type

	// configure the gradient stops
	XLGradientStops stops = fills[ testGradientIndex ].stops();

	// first XLGradientStop
	XLStyleIndex stopIndex = stops.create();
	XLDataBarColor stopColor = stops[ stopIndex ].color();
	stops[ stopIndex ].setPosition(0.1);
	stopColor.setRgb( green );
	// stopColor.setTint( 1.01 );
	// stopColor.setAutomatic();
	// stopColor.setIndexed( 77 );
	// stopColor.setTheme( 79 );

	// second XLGradientStop
	stopIndex = stops.create(stops[stopIndex]); // create another stop using previous stop as template
	stopColor.set( yellow );
	stops[ stopIndex ].setPosition(0.5);

	cellFormats[ newCellStyle ].setFillIndex( testGradientIndex );

	myCellRange.setFormat( newCellStyle ); // assign the new format to the full range of cells

	// ===== BEGIN cell borders demo
	XLStyleIndex borderFormat = 0; // scope declaration
	{
		// create a new cell format based on the current C3 format & assign it back to C3
		XLStyleIndex borderedCellFormat = cellFormats.create( cellFormats[ wks.cell("C7").cellFormat() ] );
		wks.cell("C7").setCellFormat( borderedCellFormat );
		wks.cell("C7") = "borders demo";

		// ===== Create a new border format style & assign it to the new cell format
		borderFormat = borders.create();
		// NOTE: the new border format uses a default (empty) borders object to demonstrate OpenXLSX ordered inserts.
		//       Using an existing border style as a template would copy the OpenXLSX default border style which
		//       already contains elements in the correct sequence, so that inserts could not be demonstrated.
		cellFormats[ borderedCellFormat ].setBorderIndex( borderFormat );

		borders[ borderFormat ].setDiagonalUp( false );    // setting this to true will apply the diagonal line style
		borders[ borderFormat ].setDiagonalDown( true );   //  - both diagonals can be applied simultaneously, but only with the same style
		borders[ borderFormat ].setOutline( true );        // not sure if this is needed at all

		// ===== Insert lines in a "wrong" sequence that MS Office would refuse to accept, to demonstrate XLStyles ordered inserting capability
		borders[ borderFormat ].setHorizontal( XLLineStyleHair,   XLColor( "ff446688" ) );
		borders[ borderFormat ].setVertical  ( XLLineStyleDouble, XLColor( "ff443322" ) );
		borders[ borderFormat ].setDiagonal  ( XLLineStyleThick,  XLColor( "ff332222" ), -0.25 );
		borders[ borderFormat ].setBottom    ( XLLineStyleDotted, XLColor( "ff222222" ), 0.25 );
		borders[ borderFormat ].setTop       ( XLLineStyleDashed, XLColor( "ff114444" ) );
		borders[ borderFormat ].setRight     ( XLLineStyleMedium, XLColor( "ff113333" ) );
		borders[ borderFormat ].setLeft      ( XLLineStyleThin,   XLColor( "ff112222" ) );
		std::cout << "#1 borders[ " << borderFormat << " ] summary: " << borders[ borderFormat ].summary() << std::endl;

		// test additional line styles with setter function:
		borders[ borderFormat ].setLine      ( XLLineVertical, XLLineStyleSlantDashDot,     XLColor( XLDefaultFontColor ) );
		borders[ borderFormat ].setLine      ( XLLineDiagonal, XLLineStyleMediumDashDotDot, XLColor( XLDefaultFontColor ), -1.0 );
		borders[ borderFormat ].setLine      ( XLLineBottom,   XLLineStyleDashDotDot,       XLColor( XLDefaultFontColor ), -0.1 );
		borders[ borderFormat ].setLine      ( XLLineTop,      XLLineStyleMediumDashDot,    green );
		borders[ borderFormat ].setLine      ( XLLineRight,    XLLineStyleDashDot,          XLColor( XLDefaultFontColor ), 1.0 );
		borders[ borderFormat ].setLine      ( XLLineLeft,     XLLineStyleMediumDashed,     XLColor( XLDefaultFontColor ) );
		std::cout << "#2 borders[ " << borderFormat << " ] summary: " << borders[ borderFormat ].summary() << std::endl;

		// test direct access to line properties (read: the XLDataBarColor)
		XLLine leftLine = borders[ borderFormat ].left();
		XLDataBarColor leftLineColor = leftLine.color();
		leftLineColor.setRgb( blue );
		leftLineColor.setTint( -0.2 );
		leftLineColor.setAutomatic();
		leftLineColor.setIndexed( 606 );
		leftLineColor.setTheme( 707 );
		std::cout << "#2 borders, leftLine summary: " << leftLine.summary() << std::endl;
		// unset some properties
		leftLineColor.setTint( -0.1 );
		leftLineColor.setAutomatic( false );
		leftLineColor.setIndexed( 0 );
		leftLineColor.setTheme( XLDeleteProperty );
		// NOTE: XLDeleteProperty can be used to remove a property from XML where a value of 0 would not accomplish the same
		//       formatting. Currently, only XLDataBarColor::setTheme supports this functionality. More can be added as properties
		//       are discovered that have the same problem (theme="" still overrides e.g. the line color)
		std::cout << "#2 borders, leftLineColor summary: " << leftLineColor.summary() << std::endl;
	}
	// ===== END cell borders demo ===== //

	// enable testBasics = true to create/modify at least one entry in each known styles array
	// disable testBasics = false to stop the demo execution here and save the document
	bool testBasics = false;
	// CAUTION: setting testBasics = true will cause MS Excel to complain about document errors
	//          the purpose of this scope is to prove desired modification of the underlying XML
	//          tags for direct inspection with unzip + xmllint + text editor.
	//          The testBasics scope makes no attempt to generate meaningful or consistent style information.

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
		fonts[ nodeIndex ].setScheme(XLFontSchemeMajor);
			// fonts[ nodeIndex ].setScheme(XLFontSchemeMinor);
			// fonts[ nodeIndex ].setScheme(XLFontSchemeNone);
			// fonts[ nodeIndex ].setScheme(XLFontSchemeInvalid);
		fonts[ nodeIndex ].setVertAlign(XLSubscript);
			// fonts[ nodeIndex ].setVertAlign(XLBaseline);
			// fonts[ nodeIndex ].setVertAlign(XLSuperscript);
			// fonts[ nodeIndex ].setVertAlign(XLVerticalAlignRunInvalid);
		fonts[ nodeIndex ].setOutline();
		fonts[ nodeIndex ].setShadow();
		fonts[ nodeIndex ].setCondense();
		fonts[ nodeIndex ].setExtend();
		std::cout << "fonts[ " << nodeIndex << " ] summary: " << fonts[ nodeIndex ].summary() << std::endl;

		// XLFill -> XLPatternFill test:
		nodeIndex = ( createAll ? fills.create() : fills.count() - 1 );
		fills[ nodeIndex ].setFillType( XLPatternFill );
		fills[ nodeIndex ].setPatternType( XLPatternSolid );
		// fills[ nodeIndex ].setPatternType( XLPatternTypeInvalid );
		fills[ nodeIndex ].setColor          ( XLColor( XLDefaultPatternFgColor ) );
		fills[ nodeIndex ].setBackgroundColor( XLColor( XLDefaultPatternBgColor ) );

		// output summary of previously created style with a gradient fill
		std::cout << "fills[ " << testGradientIndex << " ] summary: " << fills[ testGradientIndex ].summary() << std::endl;

		// output summary of newly created style
		std::cout << "fills[ " << nodeIndex << " ] summary: " << fills[ nodeIndex ].summary() << std::endl;


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


	XLStyleIndex mergedCellFormat = cellFormats.create( cellFormats[ wks.cell("E4").cellFormat() ] );
	cellFormats[ mergedCellFormat ].alignment(XLCreateIfMissing).setHorizontal(XLAlignCenter);
	cellFormats[ mergedCellFormat ].alignment(XLCreateIfMissing).setVertical(XLAlignTop);
	cellFormats[ mergedCellFormat ].alignment(XLCreateIfMissing).setTextRotation(5);
	cellFormats[ mergedCellFormat ].alignment(XLCreateIfMissing).setWrapText(true);

	wks.cell("E4") = "merged red range #1\n - hidden cell values are intact!";
   wks.mergeCells("E4:G7");                     // merge cells without deletion of contents

	// How to get the value of a merged range that a cell is a part of (by getting the top left cell first)
	std::cout << "value of cell F6 is " << wks.cell("F6") << std::endl;    // empty
	XLMergeCells merges = wks.merges();                                // get the worksheet's collection of merged ranges
	int32_t mergeIndexForF6 = merges.findMergeByCell("F6");            // determine if F6 is contained within any merge
	if( mergeIndexForF6 != -1 ) {                                          // if cell is found at mergeIndexForF6
		std::string mergeRef = merges[ mergeIndexForF6 ];                       // get the actual merge reference
		size_t pos = mergeRef.find_first_of(':');                               // 
		XLCellReference topLeftRef(mergeRef.substr(0, pos));                    // extract from the merge the top left cell

		// output information about the merge
		std::cout << "mergeIndexForF6 is " << mergeIndexForF6 << " , this merge reference is " << mergeRef << " and the top left cell is " << topLeftRef.address() << std::endl;
		// output the value of the merge's top left cell
		std::cout << "value of cell F6 by merge is \"" << wks.cell(topLeftRef) << "\"" << std::endl;
	}
	XLStyleIndex mergedCellFormatWithBorder = cellFormats.create( cellFormats[ mergedCellFormat ] );
	if( false ) // enable this if you want to experiment with the borderFormat
		cellFormats[ mergedCellFormatWithBorder ].setBorderIndex( borderFormat );
	wks.cell("E4").setCellFormat( mergedCellFormatWithBorder );
	wks.cell("J5") = "merged red range #2\n - hidden cell contents have been deleted!";
   wks.mergeCells("J5:L8", XLEmptyHiddenCells); // merge cells    with deletion of contents
	wks.cell("J5").setCellFormat( mergedCellFormat );

	
	// ===== BEGIN conditional formatting demo
	{
		XLDiffCellFormats & diffCellFormats = doc.styles().diffCellFormats(); // get a handle on differential cell formats

		// Create some differential cell formats that can be used for conditional formatting
		XLStyleIndex lowStyle       = diffCellFormats.create();
		XLStyleIndex justRightStyle = diffCellFormats.create();
		XLStyleIndex highStyle      = diffCellFormats.create();

		XLColor gray( "dddddddd" );
		XLColor orange( "00ffcc00" );

		diffCellFormats[ lowStyle ].font().setFontColor( blue );
		diffCellFormats[ lowStyle ].fill().setPatternType( XLPatternSolid );
		diffCellFormats[ lowStyle ].fill().setColor( gray );
		diffCellFormats[ lowStyle ].fill().setBackgroundColor( gray );

		diffCellFormats[ justRightStyle ].font().setFontColor( green );
		diffCellFormats[ justRightStyle ].fill().setPatternType( XLPatternNone );
		// diffCellFormats[ justRightStyle ].fill().setColor( yellow );

		diffCellFormats[ highStyle ].font().setFontColor( red );
		diffCellFormats[ highStyle ].fill().setPatternType( XLPatternSolid );
		diffCellFormats[ highStyle ].fill().setColor( orange );

		// Now create some rules applying conditional formatting to a range
		XLConditionalFormats fmts = wks.conditionalFormats();
		size_t newFmtIndex = fmts.create(); // will be 0 for this demo
		XLConditionalFormat fmt = fmts[ newFmtIndex ];
		fmt.setSqref("B13:B19");
		XLCfRules fmtRules = fmt.cfRules();
		// make 3 cfRule entries
		fmtRules.create();
		fmtRules.create();
		fmtRules.create();
		fmtRules[ 0 ].setType( XLCfType::CellIs );
		fmtRules[ 0 ].setOperator( XLCfOperator::LessThanOrEqual );
		fmtRules[ 0 ].setFormula("4");
		fmtRules[ 0 ].setDxfId(lowStyle);
		fmtRules[ 1 ].setType( XLCfType::CellIs );
		fmtRules[ 1 ].setOperator( XLCfOperator::Equal );
		fmtRules[ 1 ].setFormula("5");
		fmtRules[ 1 ].setDxfId(justRightStyle);
		fmtRules[ 2 ].setType( XLCfType::CellIs );
		fmtRules[ 2 ].setOperator( XLCfOperator::GreaterThanOrEqual );
		fmtRules[ 2 ].setFormula("6");
		fmtRules[ 2 ].setDxfId(highStyle);

		// Manipulating rule priorities manually:
		//                 priorities: 1, 2, 3
		fmtRules.setPriority(1, 1); // 2, 1, 4 // "mess up" the sequence
		fmtRules.setPriority(0, 1); // 1, 2, 5 // insert an existing priority at index 0
		fmtRules.setPriority(2, 3); // 1, 2, 3 // re-assign the priority 3
		fmtRules.setPriority(1, 1); // 2, 1, 4 // "mess up" the sequence
		fmtRules.setPriority(2, 3); // 2, 1, 3 // re-assign the priority 3
		fmtRules.setPriority(1, 4); // 2, 4, 3 // temp-assign priority 4 to index 1
		fmtRules.setPriority(0, 1); // 1, 4, 3 // re-assign priority 1 to index 0
		fmtRules.setPriority(1, 2); // 1, 2, 3 // re-assign priority 2 to index 1
		fmtRules.setPriority(1, 1); // 2, 1, 4 // "mess up" the sequence

		// renumber priorities while retaining the sequence:
		fmtRules.renumberPriorities(10); // 20, 10, 30 // renumber priorities with higher increment
		fmtRules.renumberPriorities(21845); // 21845, 43690, 65535 - max possible priority value used
		// not allowed:
		// fmtRules.create(); // triggers an error because no priority value can be assigned anymore
		// fmtRules.renumberPriorities(21846); // triggers an error because 3*21846 exceeds uint16_t range
		fmtRules.renumberPriorities(); // 2, 1, 3 // renumber priorities while retaining the sequence

		std::cout << std::endl << "conditionalFormatting entries summary: " << wks.conditionalFormats().summary() << std::endl;

		// Finally, set some exemplary values in the range
		wks.cell("B12") = "conditional formatting example";
		wks.cell("B12").setCellFormat( boldDefault );
		int cellValue = 1;
		for( XLCellAssignable c : wks.range("B13:B19") ) c = cellValue++;
	}
	// ===== END conditional formatting demo

	doc.saveAs("./Demo10.xlsx", XLForceOverwrite);
	doc.close();

	return 0;
}
