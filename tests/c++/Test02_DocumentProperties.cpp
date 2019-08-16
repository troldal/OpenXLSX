//
// Created by Troldal on 2019-01-13.
//

#include "catch.hpp"
#include <fstream>
#include <OpenXLSX.h>

#include <iostream>
#include <iomanip>
#include <ctime>
#include <sstream>

using namespace OpenXLSX;

/**
 * @brief
 *
 * @details
 */
TEST_CASE("C++ Interface Test 02: Testing of Document Properties") {

XLDocument doc;
std::string file = "./TestDocumentProperties.xlsx";
doc.
OpenDocument(file);

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02A: GetProperty") {
REQUIRE(doc
.
GetProperty(XLProperty::Title)
.
empty()
);
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02B: Test SetProperty - Title") {
doc.
SetProperty(XLProperty::Title,
"TitleTest");
doc.
SaveDocument();
doc.
CloseDocument();
doc.
OpenDocument(file);
REQUIRE(doc
.
GetProperty(XLProperty::Title)
== "TitleTest");
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02C: Test SetProperty - Title (Duplicate)") {
doc.
SetProperty(XLProperty::Title,
"TitleTest");
doc.
SaveDocument();
doc.
CloseDocument();
doc.
OpenDocument(file);
REQUIRE(doc
.
GetProperty(XLProperty::Title)
== "TitleTest");
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02D: Test SetProperty - Subject") {
doc.
SetProperty(XLProperty::Subject,
"SubjectTest");
doc.
SaveDocument();
doc.
CloseDocument();
doc.
OpenDocument(file);
REQUIRE(doc
.
GetProperty(XLProperty::Subject)
== "SubjectTest");
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02E: Test SetProperty - Creator") {
doc.
SetProperty(XLProperty::Creator,
"CreatorTest");
doc.
SaveDocument();
doc.
CloseDocument();
doc.
OpenDocument(file);
REQUIRE(doc
.
GetProperty(XLProperty::Creator)
== "CreatorTest");
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02F: Test SetProperty - Keywords") {
doc.
SetProperty(XLProperty::Keywords,
"A, B, C");
doc.
SaveDocument();
doc.
CloseDocument();
doc.
OpenDocument(file);
REQUIRE(doc
.
GetProperty(XLProperty::Keywords)
== "A, B, C");
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02G: Test SetProperty - Description") {
doc.
SetProperty(XLProperty::Description,
"DescriptionTest");
doc.
SaveDocument();
doc.
CloseDocument();
doc.
OpenDocument(file);
REQUIRE(doc
.
GetProperty(XLProperty::Description)
== "DescriptionTest");
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02H: Test SetProperty - LastModifiedBy") {
doc.
SetProperty(XLProperty::LastModifiedBy,
"LastModifiedByTest");
doc.
SaveDocument();
doc.
CloseDocument();
doc.
OpenDocument(file);
REQUIRE(doc
.
GetProperty(XLProperty::LastModifiedBy)
== "LastModifiedByTest");
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02I: Test SetProperty - LastPrinted") {
std::time_t t = std::time(nullptr);
std::tm tm = *std::gmtime(&t);

std::stringstream ss;
ss <<
std::put_time(& tm,
"%FT%TZ");
auto datetime = ss.str();

doc.
SetProperty(XLProperty::LastPrinted, datetime
);
doc.
SaveDocument();
doc.
CloseDocument();
doc.
OpenDocument(file);
REQUIRE(doc
.
GetProperty(XLProperty::LastPrinted)
== datetime);
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02J: Test SetProperty - CreationDate") {
std::time_t t = std::time(nullptr);
std::tm tm = *std::gmtime(&t);

std::stringstream ss;
ss <<
std::put_time(& tm,
"%FT%TZ");
auto datetime = ss.str();

doc.
SetProperty(XLProperty::CreationDate, datetime
);
doc.
SaveDocument();
doc.
CloseDocument();
doc.
OpenDocument(file);
REQUIRE(doc
.
GetProperty(XLProperty::CreationDate)
== datetime);
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02K: Test SetProperty - ModificationDate") {
std::time_t t = std::time(nullptr);
std::tm tm = *std::gmtime(&t);

std::stringstream ss;
ss <<
std::put_time(& tm,
"%FT%TZ");
auto datetime = ss.str();

doc.
SetProperty(XLProperty::ModificationDate, datetime
);
doc.
SaveDocument();
doc.
CloseDocument();
doc.
OpenDocument(file);
REQUIRE(doc
.
GetProperty(XLProperty::ModificationDate)
== datetime);
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02L: Test SetProperty - Category") {
doc.
SetProperty(XLProperty::Category,
"CategoryTest");
doc.
SaveDocument();
doc.
CloseDocument();
doc.
OpenDocument(file);
REQUIRE(doc
.
GetProperty(XLProperty::Category)
== "CategoryTest");
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02M: Test SetProperty - Application") {
doc.
SetProperty(XLProperty::Application,
"ApplicationTest");
doc.
SaveDocument();
doc.
CloseDocument();
doc.
OpenDocument(file);
REQUIRE(doc
.
GetProperty(XLProperty::Application)
== "ApplicationTest");
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02N: Test SetProperty - DocSecurity") {
doc.
SetProperty(XLProperty::DocSecurity,
"4");
doc.
SaveDocument();
doc.
CloseDocument();
doc.
OpenDocument(file);
REQUIRE(doc
.
GetProperty(XLProperty::DocSecurity)
== "4");
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02O: Test SetProperty - ScaleCrop") {
doc.
SetProperty(XLProperty::ScaleCrop,
"false");
doc.
SaveDocument();
doc.
CloseDocument();
doc.
OpenDocument(file);
REQUIRE(doc
.
GetProperty(XLProperty::ScaleCrop)
== "false");
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02P: Test SetProperty - Manager") {
doc.
SetProperty(XLProperty::Manager,
"ManagerTest");
doc.
SaveDocument();
doc.
CloseDocument();
doc.
OpenDocument(file);
REQUIRE(doc
.
GetProperty(XLProperty::Manager)
== "ManagerTest");
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02Q: Test SetProperty - Company") {
doc.
SetProperty(XLProperty::Company,
"CompanyTest");
doc.
SaveDocument();
doc.
CloseDocument();
doc.
OpenDocument(file);
REQUIRE(doc
.
GetProperty(XLProperty::Company)
== "CompanyTest");
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02R: Test SetProperty - LinksUpToDate") {
doc.
SetProperty(XLProperty::LinksUpToDate,
"false");
doc.
SaveDocument();
doc.
CloseDocument();
doc.
OpenDocument(file);
REQUIRE(doc
.
GetProperty(XLProperty::LinksUpToDate)
== "false");
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02S: Test SetProperty - SharedDoc") {
doc.
SetProperty(XLProperty::SharedDoc,
"false");
doc.
SaveDocument();
doc.
CloseDocument();
doc.
OpenDocument(file);
REQUIRE(doc
.
GetProperty(XLProperty::SharedDoc)
== "false");
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02T: Test SetProperty - HyperlinkBase") {
doc.
SetProperty(XLProperty::HyperlinkBase,
"HyperlinkBaseTest");
doc.
SaveDocument();
doc.
CloseDocument();
doc.
OpenDocument(file);
REQUIRE(doc
.
GetProperty(XLProperty::HyperlinkBase)
== "HyperlinkBaseTest");
}

/**
* @brief
*
* @details
*/
SECTION("Test 02U: Test SetProperty - HyperlinksChanged") {
doc.
SetProperty(XLProperty::HyperlinksChanged,
"false");
doc.
SaveDocument();
doc.
CloseDocument();
doc.
OpenDocument(file);
REQUIRE(doc
.
GetProperty(XLProperty::HyperlinksChanged)
== "false");
}

/**
 * @brief
 *
 * @details
 */
SECTION("Test 02V: Test SetProperty - AppVersion") {
doc.
SetProperty(XLProperty::AppVersion,
"12.0300");
doc.
SaveDocument();
doc.
CloseDocument();
doc.
OpenDocument(file);
REQUIRE(doc
.
GetProperty(XLProperty::AppVersion)
== "12.0300");
}

/**
 * @brief
 *
 * @details
 */
SECTION("DeleteProperty - Keywords") {
doc.
DeleteProperty(XLProperty::Keywords);
doc.
SaveDocument();
doc.
CloseDocument();
doc.
OpenDocument(file);
REQUIRE(doc
.
GetProperty(XLProperty::Keywords)
.
empty()
);
}

doc.
CloseDocument();

}