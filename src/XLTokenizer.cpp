//
// Created by KBA012 on 18-04-2018.
//

#include "XLTokenizer.h"

using namespace OpenXLSX;

/**
 * @details
 */
XLTokenizer::XLTokenizer()
    : m_buffer(""),
      m_token(""),
      m_delimiter(" \t\v\n\r\f")
{
    m_currPos = m_buffer.begin();
}

/**
 * @details
 */
XLTokenizer::XLTokenizer(const std::string& str, const std::string& delimiter)
    : m_buffer(str),
      m_token(""),
      m_delimiter(delimiter)
{
    m_currPos = m_buffer.begin();
}

/**
 * @details
 */
void XLTokenizer::Set(const std::string &str, const std::string &delimiter)
{
    m_buffer = str;
    m_delimiter = delimiter;
    m_currPos = m_buffer.begin();
}

/**
 * @details
 */
void XLTokenizer::SetString(const std::string &str)
{
    m_buffer = str;
    m_currPos = m_buffer.begin();
}

/**
 * @details
 */
void XLTokenizer::SetDelimiter(const std::string &delimiter)
{
    m_delimiter = delimiter;
    m_currPos = m_buffer.begin();
}

/**
 * @details
 */
std::string XLTokenizer::Next()
{
    if(m_buffer.empty()) return "";           // skip if buffer is empty
    m_token.clear();                              // reset token string
    SkipDelimiter();                      // skip leading delimiters

    // append each char to token string until it meets delimiter
    while(m_currPos != m_buffer.end() && !IsDelimiter(*m_currPos))
    {
        m_token += *m_currPos;
        ++m_currPos;
    }
    return m_token;
}

/**
 * @details
 */
void XLTokenizer::SkipDelimiter()
{
    while(m_currPos != m_buffer.end() && IsDelimiter(*m_currPos))
        ++m_currPos;
}

/**
 * @details
 */
bool XLTokenizer::IsDelimiter(char c)
{
    return (m_delimiter.find(c) != std::string::npos);
}

/**
 * @details
 */
std::vector<std::string> XLTokenizer::Split()
{
    std::vector<std::string> tokens;
    std::string token;
    while(!(token = Next()).empty())
    {
        tokens.push_back(token);
    }

    return tokens;
}
