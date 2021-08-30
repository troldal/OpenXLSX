//
// Created by Kenneth Balslev on 28/08/2021.
//

#include "XLDateTime.hpp"
#include "XLException.hpp"
#include <iostream>

namespace {

    bool isLeapYear(int year) {

        if (year == 1900) return true;
        if (year % 400 == 0 || (year % 4 == 0 && year % 100 != 0))
            return true;
        return false;
    }

    int daysInMonth(int month, int year) {
        switch (month) {
            case 1:
                return 31;
            case 2:
                return (isLeapYear(year) ? 29 : 28);
            case 3:
                return 31;
            case 4:
                return 30;
            case 5:
                return 31;
            case 6:
                return 30;
            case 7:
                return 31;
            case 8:
                return 31;
            case 9:
                return 30;
            case 10:
                return 31;
            case 11:
                return 30;
            case 12:
                return 31;
            default:
                return 0;
                //throw OpenXLSX::XLException("Invalid month. Cannot be higher than 12");
        }
    }

} // namespace

namespace OpenXLSX
{
    XLDateTime::XLDateTime() = default;

    XLDateTime::XLDateTime(double serial) : m_serial(serial) {}

    XLDateTime::XLDateTime(const tm& timepoint) {

        m_serial = 1;

        for (int i = 0; i < timepoint.tm_year; ++i) {
            m_serial += (isLeapYear(1900 + i) ? 366 : 365);
        }

        for (int i = 0; i < timepoint.tm_mon; ++i) {
            m_serial += daysInMonth(i + 1, timepoint.tm_year + 1900);
        }

        m_serial += timepoint.tm_mday;

        uint32_t seconds = timepoint.tm_hour * 3600 + timepoint.tm_min * 60 + timepoint.tm_sec;
        m_serial += seconds / 86400.0;
    }

    XLDateTime::XLDateTime(const XLDateTime& other) = default;

    XLDateTime::XLDateTime(XLDateTime&& other) noexcept = default;

    XLDateTime& XLDateTime::operator=(const XLDateTime& other) = default;

    XLDateTime& XLDateTime::operator=(XLDateTime&& other) noexcept = default;

    double XLDateTime::serial() const
    {
        return m_serial;
    }

    std::tm XLDateTime::timepoint() const
    {
        std::tm tp;
        tp.tm_year = 0;
        tp.tm_mon = 0;
        tp.tm_mday = 0;
        tp.tm_hour = 0;
        tp.tm_min = 0;
        tp.tm_sec = 0;
        double serial = m_serial;

        while (true) {
            auto days = (isLeapYear(tp.tm_year + 1900) ? 366 : 365);
            if (days > serial) break;
            serial -= days;
            ++tp.tm_year;
        }

        while (true) {
            auto days = daysInMonth(tp.tm_mon + 1, 1900 + tp.tm_year);
            if (days > serial) break;
            serial -= days;
            ++tp.tm_mon;
        }

        tp.tm_mday = static_cast<int>(serial);
        serial -= tp.tm_mday;

        tp.tm_hour = static_cast<int>(serial * 24);
        serial -= (tp.tm_hour / 24.0);

        tp.tm_min = static_cast<int>(serial * 24 * 60);
        serial -= (tp.tm_min / (24.0 * 60.0));

        tp.tm_sec = static_cast<int>(serial * 24 * 60 * 60);

        std::cout << tp.tm_year  + 1900 << std::endl;
        std::cout << tp.tm_mon << std::endl;
        std::cout << tp.tm_mday << std::endl;
        std::cout << tp.tm_hour << std::endl;
        std::cout << tp.tm_min << std::endl;
        std::cout << tp.tm_sec << std::endl;

        return tp;
    }

    int XLDateTime::test(int year) const
    {
        return (isLeapYear(year) ? 366 : 365);
    }
}