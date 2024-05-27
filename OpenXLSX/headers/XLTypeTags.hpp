//
// Created by kenne on 27/05/2024.
//

#pragma once

#include <XLDateTime.hpp>
#include <cstdint>
#include <string>

namespace OpenXLSX
{

    struct IntegralTypeTag
    {
    };
    struct BooleanTypeTag
    {
    };
    struct FloatingPointTypeTag
    {
    };
    struct StringTypeTag
    {
    };
    struct DateTimeTypeTag
    {
    };
    struct CellValueProxyTypeTag
    {
    };

    template<typename T, typename = void>
    struct TypeTag
    {
    };

    template<typename T>
    struct TypeTag<T, typename std::enable_if<std::is_integral<T>::value>::type>
    {
        using type = IntegralTypeTag;
    };

    template<>
    struct TypeTag<bool>
    {
        using type = BooleanTypeTag;
    };

    template<typename T>
    struct TypeTag<T, typename std::enable_if<std::is_floating_point<T>::value>::type>
    {
        using type = FloatingPointTypeTag;
    };

    template<typename T>
    struct TypeTag<T,
                   typename std::enable_if<std::is_same<T, std::string>::value || std::is_same<T, nonstd::string_view>::value ||
                                           std::is_same<T, char*>::value || std::is_same<T, const char*>::value>::type>
    {
        using type = StringTypeTag;
    };

    template<>
    struct TypeTag<XLDateTime>
    {
        using type = DateTimeTypeTag;
    };


}    // namespace OpenXLSX