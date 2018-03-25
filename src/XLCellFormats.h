//
// Created by Kenneth Balslev on 04/06/2017.
//

#ifndef OPENXLEXE_XLCELLFORMATS_H
#define OPENXLEXE_XLCELLFORMATS_H

namespace RapidXLSX
{
    class XMLNode;

//======================================================================================================================
//========== XLCellFormats Class =======================================================================================
//======================================================================================================================

    /**
     * @brief
     */
    class XLCellFormats
    {

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief
         */
        explicit XLCellFormats();

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:

        XMLNode *m_cellFormatNode; /**< */

        unsigned int m_numberFormatId; /**< */
        unsigned int m_fontId; /**< */
        unsigned int m_fillId; /**< */
        unsigned int m_borderId; /**< */
        unsigned int m_xfId; /**< */

        bool m_applyfont; /**< */
        bool m_applyFill; /**< */
        bool m_applyBorder; /**< */

    };

}

#endif //OPENXLEXE_XLCELLFORMATS_H
