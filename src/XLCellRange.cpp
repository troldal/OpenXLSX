//
// Created by KBA012 on 15/09/2016.
//

#include <stdexcept>

#include "XLCell.h"
#include "XLCellRange.h"
#include "XLCellValue.h"
#include "XLWorksheet.h"

using namespace std;
using namespace RapidXLSX;


/**
 * @details From the two XLCellReference objects, the constructor calculates the dimensions of the range.
 * If the range exceeds the current bounds of the spreadsheet, the spreadsheet is resized to fit.
 */
XLCellRange::XLCellRange(XLWorksheet &sheet,
                         const XLCellReference &topLeft,
                         const XLCellReference &bottomRight)
    : m_parentWorksheet(&sheet),
      m_topLeft(topLeft),
      m_bottomRight(bottomRight),
      m_rowOffset(0),
      m_columnOffset(0),
      m_rows(1),
      m_columns(1),
      m_transpose(false)
{
    // Find the bounds of the Range
    auto firstColumn = m_topLeft.Column();
    auto lastColumn = m_bottomRight.Column() + 1;
    auto firstRow = m_topLeft.Row();
    auto lastRow = m_bottomRight.Row() + 1;

    // Set the dimension and offset parameters
    m_rows = lastRow - firstRow;
    m_columns = lastColumn - firstColumn;
    m_rowOffset = firstRow - 1;
    m_columnOffset = firstColumn - 1;

    // TODO: Something fishy is goung on with the +1/-1 calcs...
    // TODO: This appears to slow down Cell access significantly!
    // Ensure that all cells in Range are valid, by setting last Cell.
    //if (Sheet->Cell(lastRow-1, lastColumn-1)->CellType() == XLCellType::Empty)
    //Sheet->Cell(lastRow-1, lastColumn-1)->SetEmptyValue();
}

/**
 * @todo This is not pretty, but it works
 */
XLCellRange::XLCellRange(const XLWorksheet &sheet,
                         const XLCellReference &topLeft,
                         const XLCellReference &bottomRight)
    : XLCellRange(const_cast<XLWorksheet &>(sheet), topLeft, bottomRight)
{

}

/**
 * @details Assign (copy) the contents of one range to another.
 * @todo Currently copies only values. Consider copying styles etc. as well
 */
XLCellRange &XLCellRange::operator=(const XLCellRange &other)
{

    if (other.NumRows() != NumRows() || other.NumColumns() != NumColumns())
        throw range_error("Ranges must be of same size and shape!");

    for (unsigned long r = 1; r <= NumRows(); ++r) {
        for (unsigned int c = 1; c <= NumColumns(); ++c) {
            if (other.Cell(r, c) == nullptr) {
                Cell(r, c)->Value().SetEmpty();
            }
            else {
                Cell(r, c)->Value() = other.Cell(r, c)->Value();
            }
        }
    }

    return *this;
}

/**
 * @details Returns a pointer to the XLCell at the given coordinates.
 * @todo return a const_cast version of the const method.
 */
XLCell *XLCellRange::Cell(unsigned long row,
                          unsigned int column)
{

    return const_cast<XLCell *>(static_cast<const XLCellRange *>(this)->Cell(row, column));
}

/**
 * @details Returns a const pointer to the XLCell at the given coordinates.
 */
const XLCell *XLCellRange::Cell(unsigned long row,
                                unsigned int column) const
{

    const XLCell *result = nullptr;
    XLCellReference cellReference;

    // Check if the coordinates are inside the Range.
    if (row > NumRows() || column > NumColumns()) throw std::range_error("Cell coordinates outside of Range");

    if (!m_transpose)
        cellReference.SetRowAndColumn(row + m_rowOffset, column + m_columnOffset);
    else
        cellReference.SetRowAndColumn(column + m_rowOffset, row + m_columnOffset);

    result = &m_parentWorksheet->Cell(cellReference);

    return result;
}

/**
 * @details Returns the m_rows property.
 */
unsigned long XLCellRange::NumRows() const
{
    if (!m_transpose)
        return m_rows;
    else
        return m_columns;
}

/**
 * @details Returns the m_columns property.
 */
unsigned int XLCellRange::NumColumns() const
{
    if (!m_transpose)
        return m_columns;
    else
        return m_rows;
}

/**
 * @details Sets the m_transpose member variable to true or false, depending on the argument. Effectively, this means
 * the coordinate system of the range is changed, such that rows become columns and columns becomes rows.
 * Copying one transposed range to another transposed range will cancel the transpose; i.e. it is equivalent to a
 * normal (non-transposed) copy operation.
 */
void XLCellRange::Transpose(bool state) const
{
    m_transpose = state;
}

/**
 * @details Returns a XLCellGridIterator, pointing to the first cell (top left) in the range.
 */
XLCellIterator XLCellRange::begin()
{
    //return XLCellIterator(this);
    return XLCellIterator(*this, Cell(1, 1));
}

/**
 * @details
 */
XLCellIteratorConst XLCellRange::begin() const
{
    return XLCellIteratorConst(*this, Cell(1, 1));
}

/**
 * @details Returns a XLCellGridIterator object, pointing to the last cell (bottom right) in the range.
 */
XLCellIterator XLCellRange::end()
{
    //return XLCellIterator();
    return XLCellIterator(*this, nullptr);
}

XLCellIteratorConst XLCellRange::end() const
{
    return XLCellIteratorConst(*this, nullptr);
}

void XLCellRange::Clear()
{

    for (unsigned long row = 1; row <= m_rows; row++) {
        for (unsigned int column = 1; column <= m_columns; column++) {
            Cell(row, column)->Value().SetEmpty();
        }
    }
}
