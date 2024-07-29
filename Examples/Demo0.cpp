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
	// DONE: xmlns namespace support on XMLNode level (to be tested)
	// DONE: in _rels/.rels: strip a leading '/' (or at least ignore it) in Target path at all times
	//       (currently, for this reason, the relationships for docProps/core.xml and docProps/app.xml are duplicated with a new Target link)
	// TODO: UUID support for worksheet r:id values
	//       - [DONE] _rels/.rels <Relationship <!-- other attributes --> Id="R1089e9da7ec94cca"/>
	//          --> used for relevant xml targets in the workbook
	//       - [DONE] xl/workbook##.xml sheets::sheet node: <x:sheet name="Workspace" sheetId="1" r:id="R4b78b06d10804515"/>
	//          --> regular sheet Ids
	//       - xl/_rels/workbook##.xml.rels: <Relationship <!-- other attributes --> Id="Rca50407821924816"/>
	//          workbook relationships: xl/sharedStrings.xml, xl/styles.xml, xl/theme/theme##.xml, and all worksheets
	//          [DONE] --> regular sheet Ids for worksheets
	//          TODO --> for the other xml files, it appears the IDs here are not used elsewhere
	//       - [NOT TOUCHED BY OpenXLSX] xl/worksheets/_rels/sheet##.xml.rels <Relationship <!-- other attributes --> Id="Rc3c7515c124f4337"/> 
	//          each sheet needs a relationship entry for 
	//          - Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
	//            --> it appears the ID here is not used elsewhere
	//          - Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
	//            --> the target for a vmlDrawing is referenced by this ID in the worksheet (see below)
	//       - [NOT TOUCHED BY OpenXLSX] xl/worksheets/sheet##.xml worksheet node "<x:legacyDrawing r:id="Rc3c7515c124f4337"/>"
	//         --> refers to the ID defined in the worksheet relationships (above) no support needed atm

	// TODO: merge cells support:
	// in sheet#.xml:
	// <x:mergeCells count="10">
	// 	<x:mergeCell ref="AD2:AE2"/>
	// 	<x:mergeCell ref="AF2:AG2"/>
	// </x:mergeCells>
	// TODO: comments support: each xl/comments##.xml file refers to the corresponding xl/worksheets/sheets##.xml file with comments structured like so:
	// <x:commentList>
	// 	<x:comment ref="M5" authorId="0">
	// 		<x:text>
	// 			<x:t>Enabling this option will reduce memory usage by 
	// saving some intermediate data from memory onto the
	// computer's hard disk.</x:t>
	// 		</x:text>
	// 	</x:comment>
	// NOTE: each comments file should be listed in [Content_Types].xml:
	//       - Types::Override node: <Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
	// </x:commentList>
   // TODO: check [Content_Types].xml "support" (in that saving the file does not break anything):
	//       - Types::Default node: <Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
	//         (alternative entry to ContentType="application/xml")
	//       - Types::Default node: <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
	//       - Types::Default node: <Default Extension="dat" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"/>



	cout << "********************************************************************************\n";
	cout << "DEMO PROGRAM ###: Load a file to test git pull request #186 Ignore whitespace in sharedStrings.xml\n";
	cout << "********************************************************************************\n";

	std::cout << "nowide is " << ( nowide_status() ? "enabled" : "disabled" ) << std::endl;

	if (enable_xml_namespaces())
		std::cout << "enabled support for xml namespaces!" << std::endl;
	else {
		std::cerr << "failed to enable support for xml namespaces!" << std::endl;
		return 1;
	}

	// configure OpenXLSX to use 64 bit random relationship IDs instead of sequential IDs, to support logic of workbook22
	UseRandomIDs();

	// load example file
	XLDocument doc("./workbook22.xlsx"s);
	doc.setSavingDeclaration( XLXmlSavingDeclaration(XLXmlDefaultVersion, "utf-8", XLXmlNotStandalone) );

	std::cout << "doc.name() is " << doc.name() << std::endl;

	if( argc > 1 && std::string( argv[1] ) == "--new-worksheet" ) {
		std::string newSheetName = "test new worksheet";
		doc.workbook().addWorksheet( newSheetName );
		auto wks = doc.workbook().worksheet( newSheetName );
		std::cout << "wks.name() is " << wks.name() << std::endl;
		wks.setActive();
	}

	doc.saveAs("./workbook22-out.xlsx", XLForceOverwrite);
	doc.close();

	return 0;
}
