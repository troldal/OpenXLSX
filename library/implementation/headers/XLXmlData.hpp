//
// Created by Kenneth Balslev on 11/07/2020.
//

#ifndef OPENXLSX_XLXMLDATA_HPP
#define OPENXLSX_XLXMLDATA_HPP

// ===== Standard Library Includes ===== //
#include <cstring>
#include <memory>
#include <sstream>
#include <string>

// ===== OpenXLSX Includes ===== //
#include "XLEnums.hpp"
#include "XLXmlParser.hpp"

namespace OpenXLSX::Impl
{
    class XLDocument;

    /**
     * @brief The XLXmlData class encapsulates the properties and behaviour of the .xml files in an .xlsx file zip
     * package. Objects othe XLXmlData type are intended to be stored centrally in an XLDocument object, from where
     * they can be retrieved by other objects that encapsulates the behaviour of Excel elements, such as XLWorkbook
     * and XLWorksheet.
     */
    class XLXmlData final
    {
    public:
        // ===== PUBLIC MEMBER FUNCTIONS ===== //

        /**
         * @brief Default constructor. All member variables are default constructed. Except for
         * the raw XML data, none of the member variables can be modified after construction. Hence, objects created
         * using the default constructor can only serve as null objects and targets for the move assignemnt operator.
         */
        XLXmlData() = default;

        /**
         * @brief Constructor. This constructor creates objects with the given parameters. the xmlId and the xmlType
         * parameters have default values. These are only useful for relationship (.rels) files and the
         * [Content_Types].xml file located in the root directory of the zip package.
         * @param parentDoc A pointer to the parent XLDocument object.
         * @param xmlPath A std::string with the file path in zip package.
         * @param xmlId A std::string with the relationship ID of the file (used in the XLRelationships class)
         * @param xmlType The type of object the XML file represents, e.g. XLWorkbook or XLWorksheet.
         */
        XLXmlData(XLDocument*        parentDoc,
                  const std::string& xmlPath,
                  const std::string& xmlId   = "",
                  XLContentType      xmlType = XLContentType::Unknown);

        /**
         * @brief Default destructor. The XLXmlData does not manage any dynamically allocated resources, so a default
         * destructor will suffice.
         */
        ~XLXmlData() = default;

        /**
         * @brief Copy constructor. The m_xmlDoc data member is a XMLDocument object, which is non-copyable. Hence,
         * the XLXmlData objects have a explicitly deleted copy constructor.
         * @param other
         */
        XLXmlData(const XLXmlData& other) = delete;

        /**
         * @brief Move constructor. All data members are trivially movable. Hence an explicitly defaulted move
         * constructor is sufficient.
         * @param other
         */
        XLXmlData(XLXmlData&& other) noexcept = default;

        /**
         * @brief Copy assignment operator. The m_xmlDoc data member is a XMLDocument object, which is non-copyable.
         * Hence, the XLXmlData objects have a explicitly deleted copy assignment operator.
         */
        XLXmlData& operator=(const XLXmlData& other) = delete;

        /**
         * @brief Move assignment operator. All data members are trivially movable. Hence an explicitly defaulted move
         * constructor is sufficient.
         * @param other the XLXmlData object to be moved from.
         * @return A reference to the moved-to object.
         */
        XLXmlData& operator=(XLXmlData&& other) noexcept = default;

        /**
         * @brief Set the raw data for the underlying XML document. Being able to set the XML data directly is useful
         * when creating a new file using a XML file template. E.g., when creating a new worksheet, the XML code for
         * a minimum viable XLWorksheet object can be added using this function.
         * @param data A std::string with the raw XML text.
         */
        void setRawData(const std::string& data);

        /**
         * @brief Get the raw data for the underlying XML document. This function will retrieve the raw XML text data
         * from the underlying XMLDocument object. This will mainly be used when saving data to the .xlsx package
         * using the save function in the XLDocument class.
         * @return A std::string with the raw XML text data.
         */
        std::string getRawData() const;

        /**
         * @brief Access the parent XLDocument object.
         * @return A pointer to the parent XLDocument object.
         */
        XLDocument* getParentDoc();

        /**
         * @brief Access the parent XLDocument object.
         * @return A const pointer to the parent XLDocument object.
         */
        const XLDocument* getParentDoc() const;

        /**
         * @brief Retrieve the path of the XML data in the .xlsx zip archive.
         * @return A std::string with the path.
         */
        std::string getXmlPath() const;

        /**
         * @brief Retrieve the relationship ID of the XML file.
         * @return A std::string with the relationship ID.
         */
        std::string getXmlID() const;

        /**
         * @brief Retrieve the type represented by the XML data.
         * @return A XLContentType value representing the type.
         */
        XLContentType getXmlType() const;

        /**
         * @brief Access the underlying XMLDocument object.
         * @return A pointer to the XMLDocument object.
         */
        XMLDocument* getXmlDocument();

        /**
         * @brief Access the underlying XMLDocument object.
         * @return A const pointer to the XMLDocument object.
         */
        const XMLDocument* getXmlDocument() const;

    private:
        // ===== PRIVATE MEMBER VARIABLES ===== //

        XLDocument*         m_parentDoc {}; /**< A pointer to the parent XLDocument object. >*/
        std::string         m_xmlPath {};   /**< The path of the XML data in the .xlsx zip archive. >*/
        std::string         m_xmlID {};     /**< The relationship ID of the XML data. >*/
        XLContentType       m_xmlType {};   /**< The type represented by the XML data. >*/
        mutable XMLDocument m_xmlDoc {};    /**< The underlying XMLDocument object. >*/
    };
}    // namespace OpenXLSX::Impl

#endif    // OPENXLSX_XLXMLDATA_HPP
