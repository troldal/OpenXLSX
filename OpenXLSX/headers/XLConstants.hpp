//
// Created by Kenneth Balslev on 15/08/2021.
//

#ifndef OPENXLSX_XLCONSTANTS_HPP
#define OPENXLSX_XLCONSTANTS_HPP

namespace OpenXLSX
{
    inline constexpr uint16_t MAX_COLS = 16'384;
    inline constexpr uint32_t MAX_ROWS = 1'048'576;

    // anchoring a comment shape below these values was not possible in LibreOffice - TBC with MS Office
    inline constexpr uint16_t MAX_SHAPE_ANCHOR_COLUMN = 13067;  // column "SHO"
    inline constexpr uint32_t MAX_SHAPE_ANCHOR_ROW    = 852177;
}    // namespace OpenXLSX

#endif    // OPENXLSX_XLCONSTANTS_HPP
