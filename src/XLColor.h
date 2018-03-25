//
// Created by Kenneth Balslev on 04/06/2017.
//

#ifndef OPENXLEXE_XLCOLOR_H
#define OPENXLEXE_XLCOLOR_H

#include <string>

namespace RapidXLSX
{

//======================================================================================================================
//========== XLColor Class =============================================================================================
//======================================================================================================================

    /**
     * @brief
     */
    class XLColor
    {

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief
         * @param red
         * @param green
         * @param blue
         */
        explicit XLColor(unsigned int red = 0,
                         unsigned int green = 0,
                         unsigned int blue = 0);

        /**
         * @brief
         * @param hexCode
         */
        explicit XLColor(const std::string &hexCode);

        /**
         * @brief
         * @param other
         */
        XLColor(const XLColor &other) = default;

        /**
         * @brief
         */
        ~XLColor() = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLColor &operator=(const XLColor &other) = default;

        /**
         * @brief
         * @param red
         * @param green
         * @param blue
         */
        void SetColor(unsigned int red = 0,
                      unsigned int green = 0,
                      unsigned int blue = 0);

        /**
         * @brief
         * @param hexCode
         */
        void SetColor(const std::string &hexCode);

        /**
         * @brief
         * @return
         */
        unsigned int Red() const;

        /**
         * @brief
         * @return
         */
        unsigned int Green() const;

        /**
         * @brief
         * @return
         */
        unsigned int Blue() const;

        /**
         * @brief
         * @return
         */
        std::string Hex() const;

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:

        unsigned int m_red;
        unsigned int m_green;
        unsigned int m_blue;

    };

}


#endif //OPENXLEXE_XLCOLOR_H
