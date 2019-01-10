//
// Created by KBA012 on 18-04-2018.
//

#include "XLTokenizer_Impl.h"
#include <algorithm>
#include <cstring>

using namespace OpenXLSX::Impl;
using namespace std;

/**
 * @details
 */
XLToken::XLToken(const std::string &token) : m_token(token)
{
}

/**
 * @details
 */
const std::string &XLToken::AsString() const
{
    return m_token;
}

/**
 * @details
 */
long long int XLToken::AsInteger() const
{
    return stoll(m_token);
}

/**
 * @details
 */
long double XLToken::AsFloat() const
{
    string temp = m_token;
    replace(temp.begin(), temp.end(), ',', '.');
    return stold(temp);
}

/**
 * @details
 */
bool XLToken::AsBoolean() const
{
    if (m_token == "FALSE") return false;
    else return true;
}

/**
 * @details
 */
bool XLToken::IsString() const
{
    if (IsBoolean() || IsFloat() || IsInteger()) return false;
    return true;
}

/**
 * @details
 */
bool XLToken::IsInteger() const
{
    for (auto &iter : m_token) {
        if (iter == *m_token.begin() && iter == '-') continue;
        if (!isdigit(iter)) return false;
    }

    return true;
}

/**
 * @details
 */
bool XLToken::IsFloat() const
{
    int numDigit = 0;
    int numDecimal = 0;
    int numOther = 0;

    for (auto &iter : m_token) {
        if (iter == *m_token.begin() && iter == '-') continue;
        if (isdigit(iter)) {
            numDigit++;
            continue;
        }

        if (iter == '.' || iter == ',') {
            numDecimal++;
            continue;
        }

        numOther++;
    }

    if (numDecimal == 1 && numOther == 0 && numDigit > 0) return true;

    return false;
}

/**
 * @details
 */
bool XLToken::IsBoolean() const
{
    string strTrue = "TRUE";
    string strFalse = "FALSE";

    bool isTrue = equal(m_token.begin(), m_token.end(), strTrue.begin(), [](int c1, int c2){return toupper(c1) == toupper(c2);});
    bool isFalse = equal(m_token.begin(), m_token.end(), strFalse.begin(), [](int c1, int c2){return toupper(c1) == toupper(c2);});

    if (isTrue || isFalse) return true;
    else return false;
}

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
XLToken XLTokenizer::Next()
{
    if(m_buffer.empty()) return XLToken("");           // skip if buffer is empty
    m_token.clear();                              // reset token string
    SkipDelimiter();                      // skip leading delimiters

    // append each char to token string until it meets delimiter
    while(m_currPos != m_buffer.end() && !IsDelimiter(*m_currPos))
    {
        m_token += *m_currPos;
        ++m_currPos;
    }
    return XLToken(m_token);
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
vector<XLToken> XLTokenizer::Split()
{
    std::vector<XLToken> tokens;
    std::string token;
    while(!(token = Next().AsString()).empty())
    {
        tokens.push_back(XLToken(token));
    }

    return tokens;
}
