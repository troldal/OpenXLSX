//
// Created by Troldal on 2019-01-13.
//

#include "catch.hpp"
#include <fstream>

#include <iostream>
#include <iomanip>
#include <ctime>
#include <sstream>

extern "C" {
#include <XL_Document.h>
}

/**
 * @brief
 *
 * @details
 */
TEST_CASE("C Interface Test 02: Testing of Document Properties") {

    std::string file = "./TestDocumentProperties.xlsx";
    XL_Document* doc = XL_OpenDocument(file.c_str());

    /**
     * @brief
     *
     * @details
     */
    SECTION("Test 02A: GetDocumentProperty") {
        REQUIRE(std::string(XL_GetDocumentProperty(doc, Title)) == "");
    }

        /**
         * @brief
         *
         * @details
         */
    SECTION("Test 02B: Test SetDocumentProperty - Title") {
        XL_SetDocumentProperty(doc, Title, "TitleTest");
        XL_SaveDocument(doc);
        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);

        doc = XL_OpenDocument(file.c_str());
        REQUIRE(std::string(XL_GetDocumentProperty(doc, Title)) == "TitleTest");
    }

        /**
         * @brief
         *
         * @details
         */
    SECTION("Test 02C: Test SetDocumentProperty - Title (Duplicate)") {
        XL_SetDocumentProperty(doc, Title, "TitleTest");
        XL_SaveDocument(doc);
        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);

        doc = XL_OpenDocument(file.c_str());
        REQUIRE(std::string(XL_GetDocumentProperty(doc, Title)) == "TitleTest");
    }

        /**
         * @brief
         */
    SECTION("Test 02D: Test SetDocumentProperty - Subject") {
        XL_SetDocumentProperty(doc, Subject, "SubjectTest");
        XL_SaveDocument(doc);
        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);

        doc = XL_OpenDocument(file.c_str());
        REQUIRE(std::string(XL_GetDocumentProperty(doc, Subject)) == "SubjectTest");
    }

        /**
         * @brief
         */
    SECTION("Test 02E: Test SetDocumentProperty - Creator") {
        XL_SetDocumentProperty(doc, Creator, "CreatorTest");
        XL_SaveDocument(doc);
        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);

        doc = XL_OpenDocument(file.c_str());
        REQUIRE(std::string(XL_GetDocumentProperty(doc, Creator)) == "CreatorTest");
    }

        /**
         * @brief
         */
    SECTION("Test 02F: Test SetDocumentProperty - Keywords") {
        XL_SetDocumentProperty(doc, Keywords, "A, B, C");
        XL_SaveDocument(doc);
        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);

        doc = XL_OpenDocument(file.c_str());
        REQUIRE(std::string(XL_GetDocumentProperty(doc, Keywords)) == "A, B, C");
    }

        /**
         * @brief
         */
    SECTION("Test 02G: Test SetDocumentProperty - Description") {
        XL_SetDocumentProperty(doc, Description, "DescriptionTest");
        XL_SaveDocument(doc);
        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);

        doc = XL_OpenDocument(file.c_str());
        REQUIRE(std::string(XL_GetDocumentProperty(doc, Description)) == "DescriptionTest");
    }

        /**
         * @brief
         */
    SECTION("Test 02H: Test SetDocumentProperty - LastModifiedBy") {
        XL_SetDocumentProperty(doc, LastModifiedBy, "LastModifiedByTest");
        XL_SaveDocument(doc);
        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);

        doc = XL_OpenDocument(file.c_str());
        REQUIRE(std::string(XL_GetDocumentProperty(doc, LastModifiedBy)) == "LastModifiedByTest");
    }

        /**
         * @brief
         */
    SECTION("Test 02I: Test SetDocumentProperty - LastPrinted") {
        std::time_t t  = std::time(nullptr);
        std::tm     tm = *std::gmtime(&t);

        std::stringstream ss;
        ss << std::put_time(&tm, "%FT%TZ");
        auto datetime = ss.str();

        XL_SetDocumentProperty(doc, LastPrinted, datetime.c_str());
        XL_SaveDocument(doc);
        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);

        doc = XL_OpenDocument(file.c_str());
        REQUIRE(std::string(XL_GetDocumentProperty(doc, LastPrinted)) == datetime.c_str());
    }

        /**
         * @brief
         */
    SECTION("Test 02J: Test SetDocumentProperty - CreationDate") {
        std::time_t t  = std::time(nullptr);
        std::tm     tm = *std::gmtime(&t);

        std::stringstream ss;
        ss << std::put_time(&tm, "%FT%TZ");
        auto datetime = ss.str();

        XL_SetDocumentProperty(doc, CreationDate, datetime.c_str());
        XL_SaveDocument(doc);
        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);

        doc = XL_OpenDocument(file.c_str());
        REQUIRE(std::string(XL_GetDocumentProperty(doc, CreationDate)) == datetime.c_str());
    }

        /**
         * @brief
         */
    SECTION("Test 02K: Test SetDocumentProperty - ModificationDate") {
        std::time_t t  = std::time(nullptr);
        std::tm     tm = *std::gmtime(&t);

        std::stringstream ss;
        ss << std::put_time(&tm, "%FT%TZ");
        auto datetime = ss.str();

        XL_SetDocumentProperty(doc, ModificationDate, datetime.c_str());
        XL_SaveDocument(doc);
        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);

        doc = XL_OpenDocument(file.c_str());
        REQUIRE(std::string(XL_GetDocumentProperty(doc, ModificationDate)) == datetime.c_str());
    }

        /**
         * @brief
         */
    SECTION("Test 02L: Test SetDocumentProperty - Category") {
        XL_SetDocumentProperty(doc, Category, "CategoryTest");
        XL_SaveDocument(doc);
        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);

        doc = XL_OpenDocument(file.c_str());
        REQUIRE(std::string(XL_GetDocumentProperty(doc, Category)) == "CategoryTest");
    }

        /**
         * @brief
         */
    SECTION("Test 02M: Test SetDocumentProperty - Application") {
        XL_SetDocumentProperty(doc, Application, "ApplicationTest");
        XL_SaveDocument(doc);
        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);

        doc = XL_OpenDocument(file.c_str());
        REQUIRE(std::string(XL_GetDocumentProperty(doc, Application)) == "ApplicationTest");
    }

        /**
         * @brief
         */
    SECTION("Test 02N: Test SetDocumentProperty - DocSecurity") {
        XL_SetDocumentProperty(doc, DocSecurity, "4");
        XL_SaveDocument(doc);
        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);

        doc = XL_OpenDocument(file.c_str());
        REQUIRE(std::string(XL_GetDocumentProperty(doc, DocSecurity)) == "4");
    }

        /**
         * @brief
         */
    SECTION("Test 02O: Test SetDocumentProperty - ScaleCrop") {
        XL_SetDocumentProperty(doc, ScaleCrop, "false");
        XL_SaveDocument(doc);
        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);

        doc = XL_OpenDocument(file.c_str());
        REQUIRE(std::string(XL_GetDocumentProperty(doc, ScaleCrop)) == "false");
    }

        /**
         * @brief
         */
    SECTION("Test 02P: Test SetDocumentProperty - Manager") {
        XL_SetDocumentProperty(doc, Manager, "ManagerTest");
        XL_SaveDocument(doc);
        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);

        doc = XL_OpenDocument(file.c_str());
        REQUIRE(std::string(XL_GetDocumentProperty(doc, Manager)) == "ManagerTest");
    }

        /**
         * @brief
         */
    SECTION("Test 02Q: Test SetDocumentProperty - Company") {
        XL_SetDocumentProperty(doc, Company, "CompanyTest");
        XL_SaveDocument(doc);
        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);

        doc = XL_OpenDocument(file.c_str());
        REQUIRE(std::string(XL_GetDocumentProperty(doc, Company)) == "CompanyTest");
    }

        /**
         * @brief
         */
    SECTION("Test 02R: Test SetDocumentProperty - LinksUpToDate") {
        XL_SetDocumentProperty(doc, LinksUpToDate, "false");
        XL_SaveDocument(doc);
        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);

        doc = XL_OpenDocument(file.c_str());
        REQUIRE(std::string(XL_GetDocumentProperty(doc, LinksUpToDate)) == "false");
    }

        /**
         * @brief
         */
    SECTION("Test 02S: Test SetDocumentProperty - SharedDoc") {
        XL_SetDocumentProperty(doc, SharedDoc, "false");
        XL_SaveDocument(doc);
        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);

        doc = XL_OpenDocument(file.c_str());
        REQUIRE(std::string(XL_GetDocumentProperty(doc, SharedDoc)) == "false");
    }

        /**
         * @brief
         */
    SECTION("Test 02T: Test SetDocumentProperty - HyperlinkBase") {
        XL_SetDocumentProperty(doc, HyperlinkBase, "HyperlinkBaseTest");
        XL_SaveDocument(doc);
        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);

        doc = XL_OpenDocument(file.c_str());
        REQUIRE(std::string(XL_GetDocumentProperty(doc, HyperlinkBase)) == "HyperlinkBaseTest");
    }

        /**
         * @brief
         */
    SECTION("Test 02U: Test SetDocumentProperty - HyperlinksChanged") {
        XL_SetDocumentProperty(doc, HyperlinksChanged, "false");
        XL_SaveDocument(doc);
        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);

        doc = XL_OpenDocument(file.c_str());
        REQUIRE(std::string(XL_GetDocumentProperty(doc, HyperlinksChanged)) == "false");
    }

        /**
         * @brief
         */
    SECTION("Test 02V: Test SetDocumentProperty - AppVersion") {
        XL_SetDocumentProperty(doc, AppVersion, "12.0300");
        XL_SaveDocument(doc);
        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);

        doc = XL_OpenDocument(file.c_str());
        REQUIRE(std::string(XL_GetDocumentProperty(doc, AppVersion)) == "12.0300");
    }

        /**
         * @brief
         */
    SECTION("DeleteDocumentProperty - Keywords") {
        XL_DeleteDocumentProperty(doc, Keywords);
        XL_SaveDocument(doc);
        XL_CloseDocument(doc);
        XL_DestroyDocument(doc);

        doc = XL_OpenDocument(file.c_str());
        REQUIRE(std::string(XL_GetDocumentProperty(doc, Keywords)) == "");
    }

    XL_CloseDocument(doc);
    XL_DestroyDocument(doc);

}