//
// Created by Kenneth Balslev on 23/08/2020.
//

#ifndef OPENXLSX_XLCELLVALUEPROXY_INL
#define OPENXLSX_XLCELLVALUEPROXY_INL

/**
 * @details
 * @pre
 * @post
 */
template<typename T, typename std::enable_if<std::is_integral<T>::value, int64_t>::type*>
XLCell& XLCellValueProxy::operator=(T numberValue)
{
    if constexpr (std::is_same<T, bool>::value) {    // if bool
        setBoolean(numberValue);
    }
    else {    // if not bool
        setInteger(numberValue);
    }

    return *m_cell;
}

/**
 * @details
 * @pre
 * @post
 */
template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type*>
XLCell& XLCellValueProxy::operator=(T numberValue)
{
    setFloat(numberValue);
    return *m_cell;
}

/**
 * @details
 * @pre
 * @post
 */
template<typename T,
    typename std::enable_if<!std::is_same<T, XLCellValue>::value && !std::is_same<T, bool>::value &&
        std::is_constructible<T, const char*>::value,
                            T>::type*>
XLCell& XLCellValueProxy::operator=(T stringValue)
{
    if constexpr (std::is_same<const char*, typename std::remove_reference<typename std::remove_cv<T>::type>::type>::value)
        setString(stringValue);
    else if constexpr (std::is_same<std::string_view, T>::value)
        setString(std::string(stringValue).c_str());
    else
        setString(stringValue.c_str());
    return *m_cell;
}

/**
 * @details
 * @pre
 * @post
 */
template<typename T, typename std::enable_if<std::is_same<T, XLCellValue>::value, T>::type*>
XLCell& XLCellValueProxy::operator=(T value)
{
    switch (value.type()) {
        case XLValueType::Boolean:
            setBoolean(value.template get<bool>());
            break;
        case XLValueType::Integer:
            setInteger(value.template get<int64_t>());
            break;
        case XLValueType::Float:
            setFloat(value.template get<double>());
            break;
        case XLValueType::String:
            setString(value.template get<std::string>().c_str());
            break;
        case XLValueType::Empty:
            break;
        default:
            break;
    }

    return *m_cell;
}

/**
 * @details
 * @pre
 * @post
 */
template<typename T, typename std::enable_if<std::is_integral<T>::value, int64_t>::type*>
void XLCellValueProxy::set(T numberValue)
{
    *this = numberValue;
}

/**
 * @details
 * @pre
 * @post
 */
template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type*>
void XLCellValueProxy::set(T numberValue)
{
    *this = numberValue;
}

/**
 * @details
 * @pre
 * @post
 */
template<typename T,
    typename std::enable_if<!std::is_same<T, XLCellValue>::value && !std::is_same<T, bool>::value &&
        std::is_constructible<T, const char*>::value,
                            T>::type*>
void XLCellValueProxy::set(T stringValue)
{
    *this = stringValue;
}

/**
 * @details
 * @pre
 * @post
 */
template<typename T>
T XLCellValueProxy::get() const
{
    return getValue().get<T>();
}


#endif    // OPENXLSX_XLCELLVALUEPROXY_INL
