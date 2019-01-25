//
// Created by Troldal on 2019-01-25.
//

#ifndef OPENXLSX_XLCOLUMN_H
#define OPENXLSX_XLCOLUMN_H

namespace OpenXLSX {
    namespace Impl {
        class XLColumn;
    }

    class XLColumn
    {
    public:
        explicit XLColumn(Impl::XLColumn& column);

        XLColumn(const XLColumn& other) = default;

        XLColumn(XLColumn&& other) = default;

        virtual ~XLColumn() = default;

        XLColumn& operator=(const XLColumn& other) = default;

        XLColumn& operator=(XLColumn&& other) = default;

        /**
         * @brief Get the width of the column.
         * @return The width of the column.
         */
        float Width() const;

        /**
         * @brief Set the width of the column
         * @param width The width of the column
         */
        void SetWidth(float width);

        /**
         * @brief Is the column hidden?
         * @return The state of the column.
         */
        bool IsHidden() const;

        /**
         * @brief Set the column to be shown or hidden.
         * @param state The state of the column.
         */
        void SetHidden(bool state);

    private:
        Impl::XLColumn* m_column;

    };
}



#endif //OPENXLSX_XLCOLUMN_H
