//
// Created by Troldal on 02/09/16.
//

#ifndef OPENXL_XLSTYLES_H
#define OPENXL_XLSTYLES_H

#include "XLAbstractXMLFile.h"
#include "XLFont.h"
#include "XLSpreadsheetElement.h"
#include <map>
#include <string>

namespace RapidXLSX
{


/**
 * @brief
 */
    class XLStyles: public XLAbstractXMLFile,
                    public XLSpreadsheetElement
    {
    public:
        friend class XLFont;

        /**
         * @brief
         * @param parent
         * @param filePath
         */
        explicit XLStyles(XLDocument &parent,
                          const std::string &filePath);

        /**
         * @brief
         * @param other
         */
        XLStyles(const RapidXLSX::XLStyles &other) = delete;

        /**
         * @brief
         */
        ~XLStyles();

        /**
         * @brief
         * @return
         */
        XLStyles &operator=(const RapidXLSX::XLStyles &) = delete;

        /**
         * @brief
         * @param font
         */
        void AddFont(const XLFont &font);

        /**
         * @brief
         * @param id
         * @return
         */
        XLFont &Font(unsigned int id);

        /**
         * @brief
         * @param font
         * @return
         */
        unsigned int FontId(const std::string &font);


    protected:

        /**
         * @brief
         * @return
         */
        virtual bool ParseXMLData() override;

        //static XMLNode *FontsNode() const;

    private:

        static std::map<std::string, XLStyles> s_styles;

        static XMLNode *m_numberFormatsNode; /**<  */
        static XMLNode *m_fontsNode; /**<  */
        static XMLNode *m_fillsNode; /**<  */
        static XMLNode *m_bordersNode; /**<  */
        static XMLNode *m_cellFormatNode; /**<  */
        static XMLNode *m_cellStyleNode; /**<  */
        static XMLNode *m_colors; /**<  */

        std::vector<std::unique_ptr<XLFont>> m_fonts; /**<  */





    };

}

#endif //OPENXL_XLSTYLES_H
