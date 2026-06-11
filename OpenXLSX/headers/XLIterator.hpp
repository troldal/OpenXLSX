//
// Created by Kenneth Balslev on 22/08/2020.
//

#ifndef OPENXLSX_XLITERATOR_HPP
#define OPENXLSX_XLITERATOR_HPP

#include <memory> // std::unique_ptr

namespace OpenXLSX
{
    enum class XLIteratorDirection { Forward, Reverse };
    enum class XLIteratorLocation { Begin, End };

    template<typename T>
    class OPENXLSX_EXPORT XLReverseRange {
    public:
        /**
         * @brief Construct a reverse range from an lvalue (existing variable)
         */
        XLReverseRange(T& rr)
            : m_RR(nullptr),
              m_rr(rr)
        {}
        /**
         * @brief Construct a reverse range from an rvalue (temporary variable)
         */
        XLReverseRange(T&& rr)
            : m_RR(std::make_unique<T>(rr)),
              m_rr(*m_RR)
        {}
        typename T::reverse_iterator begin() { return m_rr.rbegin(); }
        typename T::reverse_iterator end() { return m_rr.rend(); }
    private:
        std::unique_ptr<T> m_RR;
        T& m_rr;
    };
}    // namespace OpenXLSX

#endif    // OPENXLSX_XLITERATOR_HPP
