//
// Created by Kenneth Balslev on 28/08/2021.
//

#include "XLDateTime.hpp"
#include "XLException.hpp"
#include <string>
#include <cmath>

namespace {

    /**
     * @brief
     * @param year
     * @return
     */
    bool isLeapYear(int year) {

        if (year == 1900) return true;
        if (year % 400 == 0 || (year % 4 == 0 && year % 100 != 0))
            return true;
        return false;
    }

    /**
     * @brief
     * @param month
     * @param year
     * @return
     */
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
        }
    }

    /**
     * @brief
     * @param serial
     * @return
     */
    int dayOfWeek(double serial) {
        auto day = static_cast<int32_t>(serial) % 7;
        return (day == 0 ? 6 : day - 1);
    }

} // namespace

namespace OpenXLSX
{
    /**
     * @details Conctructor. Default implementation.
     */
    XLDateTime::XLDateTime() = default;

    /**
     * @details Constructor taking an Excel date/time serial number as an argument.
     */
    XLDateTime::XLDateTime(double serial) : m_serial(serial) {
        if (serial < 1.0) throw XLDateTimeError("Excel date/time serial number is invalid (must be >= 1.0.)");
    }

    /**
     * @details Constructor taking a std::tm object as an argument.
     */
    XLDateTime::XLDateTime(const std::tm& timepoint) {

        // ===== Check validity of tm struct.
        // ===== Only year, month and day of the month are checked. Other variables are ignored.
        if (timepoint.tm_year < 0)
            throw XLDateTimeError("Invalid year. Must be >= 0.");
        if (timepoint.tm_mon < 0 || timepoint.tm_mon > 11)
            throw XLDateTimeError("Invalid month. Must be >= 0 or <= 11.");
        if (timepoint.tm_mday <= 0 || timepoint.tm_mday > daysInMonth(timepoint.tm_mon + 1, timepoint.tm_year + 1900))
            throw XLDateTimeError("Invalid day. Must be >= 1 or <= total days in the month.");

        // ===== Count the number of days for full years past 1900
        for (int i = 0; i < timepoint.tm_year; ++i) {
            m_serial += (isLeapYear(1900 + i) ? 366 : 365);
        }

        // ===== Count the number of days for full months of the last year
        for (int i = 0; i < timepoint.tm_mon; ++i) {
            m_serial += daysInMonth(i + 1, timepoint.tm_year + 1900);
        }

        // ===== Add the number of days of the month, minus one.
        // ===== (The reason for the 'minus one' is that unlike the other fields in the struct,
        // ===== tm_day represents the date of a month, whereas the other fields typically
        // ===== represents the number of whole units since the start).
        m_serial += timepoint.tm_mday - 1;

        // ===== Convert hour, minute and second to fraction of a full day.
        int32_t seconds = timepoint.tm_hour * 3600 + timepoint.tm_min * 60 + timepoint.tm_sec;
        m_serial += seconds / 86400.0;
    }

    /**
     * @details Constructor taking a unixtime format (seconds since 1/1/1970) as an argument.
     */
    XLDateTime::XLDateTime(time_t unixtime) {
        // There are 86400 seconds in a day
        // There are 25569 days between 1/1/1970 and 30/12/1899 (the epoch used by Excel)
        m_serial = static_cast<double>(unixtime) / 86400 + 25569;
    }

    /**
     * @details Copy constructor. Default implementation.
     */
    XLDateTime::XLDateTime(const XLDateTime& other) = default;

    /**
     * @details Move constructor. Default implementation.
     */
    XLDateTime::XLDateTime(XLDateTime&& other) noexcept = default;

    /**
     * @details Destructor. Default implementation.
     */
    XLDateTime::~XLDateTime() = default;

    /**
     * @details Copy assignment operator. Default implementation.
     */
    XLDateTime& XLDateTime::operator=(const XLDateTime& other) = default;

    /**
     * @details Move assignment operator. Default implementation.
     */
    XLDateTime& XLDateTime::operator=(XLDateTime&& other) noexcept = default;

    /**
     * @details
     */
    XLDateTime& XLDateTime::operator=(double serial)
    {
        XLDateTime temp(serial);
        std::swap(*this, temp);
        return *this;
    }

    /**
     * @details
     */
    XLDateTime& XLDateTime::operator=(const std::tm& timepoint)
    {
        XLDateTime temp(timepoint);
        std::swap(*this, temp);
        return *this;
    }

    /**
     * @details
     */
    XLDateTime::operator std::tm() const
    {
        return tm();
    }

    /**
     * @details Get the time point as an Excel date/time serial number.
     */
    double XLDateTime::serial() const
    {
        return m_serial;
    }

    /**
     * @details Get the time point as a std::tm object.
     */
    std::tm XLDateTime::tm() const
    {
        // ===== Create and initialize the resulting object.
        std::tm result {};
        result.tm_year = 0;
        result.tm_mon = 0;
        result.tm_mday = 0;
        result.tm_wday = 0;
        result.tm_yday = 0;
        result.tm_hour = 0;
        result.tm_min = 0;
        result.tm_sec = 0;
        result.tm_isdst = -1;
        double serial = m_serial;

        // ===== Count the number of whole years since 1900.
        while (true) {
            auto days = (isLeapYear(result.tm_year + 1900) ? 366 : 365);
            if (days > serial) break;
            serial -= days;
            ++result.tm_year;
        }

        // ===== Calculate the day of the year, and the day of the week
        result.tm_yday = static_cast<int>(serial) - 1;
        result.tm_wday = dayOfWeek(m_serial);

        // ===== Count the number of whole months in the year.
        while (true) {
            auto days = daysInMonth(result.tm_mon + 1, 1900 + result.tm_year);
            if (days > serial) break;
            serial -= days;
            ++result.tm_mon;
        }

        // ===== Calculate the number of days.
        result.tm_mday = static_cast<int>(serial);
        serial -= result.tm_mday;

        // ===== Calculate the number of hours.
        result.tm_hour = static_cast<int>(serial * 24);
        serial -= (result.tm_hour / 24.0);

        // ===== Calculate the number of minutes.
        result.tm_min = static_cast<int>(serial * 24 * 60);
        serial -= (result.tm_min / (24.0 * 60.0));

        // ===== Calculate the number of seconds.
        result.tm_sec = static_cast<int>(lround(serial * 24 * 60 * 60));

        return result;
    }

} // namespace OpenXLSX
