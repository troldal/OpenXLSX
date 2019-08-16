#include <utility>

#include <utility>

#include <utility>

//
// Created by KBA012 on 18-04-2018.
//

#include "XLTokenizer_Impl.h"
#include <algorithm>

using namespace OpenXLSX;
using namespace std;

/**
 * @details
 */
Impl::XLToken::XLToken(std::string token)
        : m_token(std::move(token)) {
}

/**
 * @details
 */
const std::string& Impl::XLToken::AsString() const {

    return m_token;
}

/**
 * @details
 */
long long int Impl::XLToken::AsInteger() const {

    return stoll(m_token);
}

/**
 * @details
 */
long double Impl::XLToken::AsFloat() const {

    string temp = m_token;
    replace(temp.begin(), temp.end(), ',', '.');
    return stold(temp);
}

/**
 * @details
 */
bool Impl::XLToken::AsBoolean() const {

    return m_token != "FALSE";
}

/**
 * @details
 */
bool Impl::XLToken::IsString() const {

    return !(IsBoolean() || IsFloat() || IsInteger());
}

/**
 * @details
 */
bool Impl::XLToken::IsInteger() const {

    for (auto& iter : m_token) {
        if (iter == *m_token.begin() && iter == '-')
            continue;
        if (isdigit(iter) == 0)
            return false;
    }

    return true;
}

/**
 * @details
 */
bool Impl::XLToken::IsFloat() const {

    int numDigit = 0;
    int numDecimal = 0;
    int numOther = 0;

    for (auto& iter : m_token) {
        if (iter == *m_token.begin() && iter == '-')
            continue;
        if (isdigit(iter) != 0) {
            numDigit++;
            continue;
        }

        if (iter == '.' || iter == ',') {
            numDecimal++;
            continue;
        }

        numOther++;
    }

    return numDecimal == 1 && numOther == 0 && numDigit > 0;

}

/**
 * @details
 */
bool Impl::XLToken::IsBoolean() const {

    string strTrue = "TRUE";
    string strFalse = "FALSE";

    bool isTrue = equal(m_token.begin(),
                        m_token.end(),
                        strTrue.begin(),
                        [](int c1, int c2) {
                            return toupper(c1) == toupper(c2);
                        });
    bool isFalse = equal(m_token.begin(),
                         m_token.end(),
                         strFalse.begin(),
                         [](int c1, int c2) {
                             return toupper(c1) == toupper(c2);
                         });

    return isTrue || isFalse;
}

/**
 * @details
 */
Impl::XLTokenizer::XLTokenizer()
        : m_buffer(""),
          m_token(""),
          m_delimiter(" \t\v\n\r\f") {

    m_currPos = m_buffer.begin();
}

/**
 * @details
 */
Impl::XLTokenizer::XLTokenizer(std::string str, std::string delimiter)
        : m_buffer(std::move(str)),
          m_token(""),
          m_delimiter(std::move(delimiter)) {

    m_currPos = m_buffer.begin();
}

/**
 * @details
 */
void Impl::XLTokenizer::Set(const std::string& str, const std::string& delimiter) {

    m_buffer = str;
    m_delimiter = delimiter;
    m_currPos = m_buffer.begin();
}

/**
 * @details
 */
void Impl::XLTokenizer::SetString(const std::string& str) {

    m_buffer = str;
    m_currPos = m_buffer.begin();
}

/**
 * @details
 */
void Impl::XLTokenizer::SetDelimiter(const std::string& delimiter) {

    m_delimiter = delimiter;
    m_currPos = m_buffer.begin();
}

/**
 * @details
 */
Impl::XLToken Impl::XLTokenizer::Next() {

    if (m_buffer.empty())
        return XLToken("");           // skip if buffer is empty
    m_token.clear();                              // reset token string
    SkipDelimiter();                      // skip leading delimiters

    // append each char to token string until it meets delimiter
    while (m_currPos != m_buffer.end() && !IsDelimiter(*m_currPos)) {
        m_token += *m_currPos;
        ++m_currPos;
    }
    return XLToken(m_token);
}

/**
 * @details
 */
void Impl::XLTokenizer::SkipDelimiter() {

    while (m_currPos != m_buffer.end() && IsDelimiter(*m_currPos))
        ++m_currPos;
}

/**
 * @details
 */
bool Impl::XLTokenizer::IsDelimiter(char c) {

    return (m_delimiter.find(c) != std::string::npos);
}

/**
 * @details
 */
vector<Impl::XLToken> Impl::XLTokenizer::Split() {

    std::vector<XLToken> tokens;
    std::string token;
    while (!(token = Next().AsString()).empty()) {
        tokens.emplace_back(token);
    }

    return tokens;
}
