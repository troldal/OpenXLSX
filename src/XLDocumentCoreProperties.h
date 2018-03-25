//
// Created by Troldal on 18/08/16.
//

#ifndef OPENXL_XLDOCCOREPROPERTIES_H
#define OPENXL_XLDOCCOREPROPERTIES_H

#include "XLAbstractXMLFile.h"
#include "XML/XMLNode.h"
#include "XLSpreadsheetElement.h"

#include <string>
#include <vector>
#include <map>

namespace RapidXLSX
{

//======================================================================================================================
//========== XLDocCoreProperties Class =================================================================================
//======================================================================================================================

    /**
     * @brief
     */
    class XLDocCoreProperties: public XLAbstractXMLFile,
                               public XLSpreadsheetElement
    {

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:


        /**
         * @brief
         * @param parent
         * @param filePath
         * @return
         */
        explicit XLDocCoreProperties(XLDocument &parent,
                                     const std::string &filePath);

        /**
         * @brief
         */
        virtual ~XLDocCoreProperties();

        /**
         * @brief
         * @param name
         * @param value
         * @return
         */
        bool SetProperty(const std::string &name,
                         const std::string &value);

        /**
         * @brief
         * @param name
         * @param value
         * @return
         */
        bool SetProperty(const std::string &name,
                         int value);

        /**
         * @brief
         * @param name
         * @param value
         * @return
         */
        bool SetProperty(const std::string &name,
                         double value);

        /**
         * @brief
         * @param name
         * @return
         */
        XMLNode &Property(const std::string &name) const;

        /**
         * @brief
         * @param name
         */
        void DeleteProperty(const std::string &name);

//----------------------------------------------------------------------------------------------------------------------
//           Protected Member Functions
//----------------------------------------------------------------------------------------------------------------------

    protected:

        /**
         * @brief
         * @return
         */
        virtual bool ParseXMLData() override;

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Functions
//----------------------------------------------------------------------------------------------------------------------

    private:
        std::map<std::string, XMLNode *> m_properties; /**< */
    };

}

#endif //OPENXL_XLDOCCOREPROPERTIES_H
