//
// Created by Troldal on 01/09/16.
//

#ifndef OPENXL_XLSHAREDSTRINGS_H
#define OPENXL_XLSHAREDSTRINGS_H

#include "XLAbstractXMLFile.h"
#include "XLSpreadsheetElement.h"

namespace RapidXLSX
{

    class XLDocument;
//======================================================================================================================
//========== XLSharedStrings Class =====================================================================================
//======================================================================================================================

/**
 * @brief This class encapsulate the Excel concept of Shared Strings. In Excel, instead of havig individual strings
 * in each cell, cells have a reference to an entry in the SharedStrings register. This results in smalle file sizes,
 * as repeated strings are referenced easily.
 * @todo Consider defining a static method for creating a new shared strings object + XML file.
 */
    class XLSharedStrings: public XLAbstractXMLFile,
                           public XLSpreadsheetElement
    {

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:


        /**
         * @brief Constructor
         * @param parent A pointer to the parent XLDocument
         * @param filePath The path to the sharedStrings.xml file
         */
        explicit XLSharedStrings(XLDocument &parent,
                                 const std::string &filePath);

        /**
         * @brief Destructor
         */
        virtual ~XLSharedStrings();

        /**
         * @brief Get the string at a given index.
         * @param index The index to look up
         * @return A std::string with the contents of the shared string
         */
        const std::string &GetString(unsigned long index) const;

        /**
         * @details Get a reference to a shared string with the given value.
         * @param str The string to look up.
         * @return A reference to the shared string.
         */
        const std::string &GetString(const std::string &str) const;

        const std::string &GetString(const std::string &str);

        /**
         * @brief Get a pointer to the XMLNode holding the shared string at a given index.
         * @param index The index to look up.
         * @return A pointer to the XMLNode holding the shared string.
         */
        const XMLNode &GetStringNode(unsigned long index) const;

        /**
         * @brief Get a pointer to the XMLNode with the given string. If more than one XMLNode exists with the lookup
         * string, the first will be returned.
         * @param str The string to look up.
         * @return A pointer to the XMLNode holding the shared string.
         */
        const XMLNode &GetStringNode(const std::string &str) const;

        /**
         * @brief
         * @param str
         * @return
         */
        long GetStringIndex(const std::string &str) const;

        /**
         * @brief
         * @param str
         * @return
         */
        bool StringExists(const std::string &str) const;

        /**
         * @brief
         * @param index
         * @return
         */
        bool StringExists(unsigned long index) const;

        /**
         * @brief Append a new string to the list of shared strings.
         * @param str The string to append.
         * @return A long int with the index of the appended string
         */
        long AppendString(const std::string &str);

        /**
         * @brief Clear the string at the given index.
         * @param index The index to clear.
         * @note There is no 'deleteString' member function, as deleting a shared string node will invalidate the
         * shared string indeces for the cells in the spreadsheet. Instead use this member functions, which clears
         * the contents of the string, but keeps the XMLNode holding the string.
         */
        void ClearString(int index);

//----------------------------------------------------------------------------------------------------------------------
//           Protected Member Functions
//----------------------------------------------------------------------------------------------------------------------

    protected:


        /**
         * @brief Parse the contnts of the underlying XML file and fills the datastructure with the data from the XML file.
         * @return true if successful; otherwise false.
         */
        bool ParseXMLData() override;

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Functions
//----------------------------------------------------------------------------------------------------------------------

    private:

        std::vector<XMLNode *> m_sharedStrings; /**< A std::vector with the XMLNodes holding the shared strings. */
        std::string m_emptyString; /**< A dummy member used for returning an empty string. */

    };

}

#endif //OPENXL_XLSHAREDSTRINGS_H
