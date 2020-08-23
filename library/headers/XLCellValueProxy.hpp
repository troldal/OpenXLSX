//
// Created by Kenneth Balslev on 23/08/2020.
//

#ifndef OPENXLSX_XLCELLVALUEPROXY_HPP
#define OPENXLSX_XLCELLVALUEPROXY_HPP

#include <type_traits>
#include <cstdint>

#include "XLCellValue.hpp"

namespace OpenXLSX {

    class XLCell;

    /**
     * @brief
     */
    class XLCellValueProxy
    {
        friend class XLCell;
        friend class XLCellValue;

    public:
        explicit XLCellValueProxy(XLCell* cell);

        ~XLCellValueProxy();

        XLCellValueProxy(const XLCellValueProxy& other);

        XLCellValueProxy& operator=(const XLCellValueProxy& other);

        template<typename T, typename std::enable_if<std::is_integral<T>::value, int64_t>::type* = nullptr>
        XLCell& operator=(T numberValue);    // NOLINT

        template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type* = nullptr>
        XLCell& operator=(T numberValue);    // NOLINT

        template<typename T,
            typename std::enable_if<!std::is_same<T, XLCellValue>::value && !std::is_same<T, bool>::value &&
                std::is_constructible<T, const char*>::value,
                                    T>::type* = nullptr>
        XLCell& operator=(T stringValue);    // NOLINT

        template<typename T, typename std::enable_if<std::is_same<T, XLCellValue>::value, T>::type* = nullptr>
        XLCell& operator=(T value);    // NOLINT

        template<typename T, typename std::enable_if<std::is_integral<T>::value, int64_t>::type* = nullptr>
        void set(T numberValue);

        template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type* = nullptr>
        void set(T numberValue);

        template<typename T,
            typename std::enable_if<!std::is_same<T, XLCellValue>::value && !std::is_same<T, bool>::value &&
                std::is_constructible<T, const char*>::value,
                                    T>::type* = nullptr>
        void set(T stringValue);

        template<typename T>
        T get() const;

        operator XLCellValue();    // NOLINT

    private:

        XLCellValueProxy(XLCellValueProxy&& other) noexcept;

        XLCellValueProxy& operator=(XLCellValueProxy&& other) noexcept;


        /**
         * @brief
         * @param numberValue
         */
        void setInteger(int64_t numberValue);

        /**
         * @brief
         * @param numberValue
         */
        void setBoolean(bool numberValue);

        /**
         * @brief
         * @param numberValue
         */
        void setFloat(double numberValue);

        /**
         * @brief
         * @param stringValue
         */
        void setString(const char* stringValue);

        /**
         * @brief Get a reference to the XLCellValue object for the cell.
         * @return A reference to an XLCellValue object.
         */
        XLCellValue getValue() const;


        XLCell* m_cell;
    };

#include "impl/XLCellValueProxy.inl"

}  // namespace OpenXLSX

#endif    // OPENXLSX_XLCELLVALUEPROXY_HPP
