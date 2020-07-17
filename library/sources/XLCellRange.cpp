//
// Created by KBA012 on 15/09/2016.
//

#include "XLCellRange.hpp"

#include "XLCell.hpp"
#include "XLWorksheet.hpp"

#include <stdexcept>

using namespace std;
using namespace OpenXLSX;

/**
 * @details From the two XLCellReference objects, the constructor calculates the dimensions of the range.
 * If the range exceeds the current bounds of the spreadsheet, the spreadsheet is resized to fit.
 */
XLCellRange::XLCellRange(XMLNode dataNode, const XLCellReference& topLeft, const XLCellReference& bottomRight)
    : m_dataNode(dataNode),
      m_topLeft(topLeft),
      m_bottomRight(bottomRight)
{}

/**
 * @details Assign (copy) the contents of one range to another.
 * @todo Currently copies only values. Consider copying styles etc. as well
 */
XLCellRange& XLCellRange::operator=(const XLCellRange& other)
{
    if (&other != this) {
        //        if (other.NumRows() != NumRows() || other.NumColumns() != NumColumns())
        //            throw range_error("Ranges must be of same size and shape!");
        //
        //        for (uint32_t r = 1; r <= NumRows(); ++r) {
        //            for (uint16_t c = 1; c <= NumColumns(); ++c) {
        //                if (other.Cell(r, c) == XLCell()) {
        //                    Cell(r, c).Value().Clear();
        //                }
        //                else {
        //                    Cell(r, c).Value() = other.Cell(r, c).Value();
        //                }
        //            }
        //        }
    }

    return *this;
}

/**
 * @details Returns the m_rows property.
 */
uint32_t XLCellRange::NumRows() const
{
    return m_bottomRight.Row() + 1 - m_topLeft.Row();
}

/**
 * @details Returns the m_columns property.
 */
uint16_t XLCellRange::NumColumns() const
{
    return m_bottomRight.Column() + 1 - m_topLeft.Column();
}

void XLCellRange::Clear()
{
    //    for (uint32_t row = 1; row <= NumRows(); row++) {
    //        for (uint16_t column = 1; column <= NumColumns(); column++) {
    //            Cell(row, column).Value().Clear();
    //        }
    //    }
}
