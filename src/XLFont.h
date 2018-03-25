//
// Created by Kenneth Balslev on 04/06/2017.
//

#ifndef OPENXLEXE_XLFONT_H
#define OPENXLEXE_XLFONT_H

#include <string>
#include <map>
#include "XLColor.h"

namespace RapidXLSX
{

    class XMLNode;

    class XLFont
    {
    public:
        friend class XLStyles;

        explicit XLFont(const std::string &name = "Cambria",
                        unsigned int size = 11,
                        const XLColor &color = XLColor(),
                        bool bold = false,
                        bool italics = false,
                        bool underline = false);

        /**
         * @brief
         * @param other
         */
        XLFont(const XLFont &other) = default;

        /**
         * @brief
         */
        ~XLFont() = default;


        XLFont &operator=(const XLFont &other) = delete;

        std::string UniqueId() const;

    private:

        static std::map<std::string, XLFont> s_fonts;

        XMLNode *m_fontNode;

        std::string m_name;
        unsigned int m_size;
        XLColor m_color;
        bool m_bold;
        bool m_italics;
        bool m_underline;

        std::string m_theme;
        std::string m_family;
        std::string m_scheme;

    };

}

#endif //OPENXLEXE_XLFONT_H
