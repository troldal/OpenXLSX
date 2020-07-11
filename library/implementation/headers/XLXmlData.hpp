//
// Created by Kenneth Balslev on 11/07/2020.
//

#ifndef OPENXLSX_XLXMLDATA_HPP
#define OPENXLSX_XLXMLDATA_HPP

#include <string>
#include <memory>
#include <cstring>

namespace OpenXLSX::Impl
{

    class XLXmlData
    {
    public:
        XLXmlData() = default;
        explicit XLXmlData(const std::string& path) : m_path(path) {}
        XLXmlData(const std::string& path, const char* data) : m_path(path), m_data(std::make_unique<char[]>(std::strlen(data))){
            std::memcpy(m_data.get(), data, std::strlen(data));
        }
        XLXmlData(const std::string& path, const void* data, size_t size) : m_path(path), m_data(std::make_unique<char[]>(size)){
            std::memcpy(m_data.get(), data, size);
        }
        ~XLXmlData() = default;
        XLXmlData(const XLXmlData& other) = delete;
        XLXmlData(XLXmlData&& other) noexcept = default;
        XLXmlData& operator=(const XLXmlData& other) = delete;
        XLXmlData& operator=(XLXmlData&& other) noexcept = default;

        inline void setData(const char* data) {
            m_data = std::make_unique<char[]>(std::strlen(data));
            std::memcpy(m_data.get(), data, std::strlen(data));
        }

        inline void setData(const void* data, size_t size) {
            m_data = std::make_unique<char[]>(size);
            std::memcpy(m_data.get(), data, size);
        }

        inline void* getData() const { return m_data.get(); }

        inline size_t getSize() const { return sizeof(m_data); }

        inline const std::string& getPath() const { return m_path; }

    private:
        std::string m_path {};
        std::unique_ptr<char[]> m_data {};
    };

}

#endif //OPENXLSX_XLXMLDATA_HPP
