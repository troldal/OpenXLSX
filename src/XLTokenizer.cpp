//
// Created by KBA012 on 18-04-2018.
//

#include "XLTokenizer.h"

using namespace OpenXLSX;

/**
 * @details
 */
XLTokenizer::XLTokenizer() : buffer(""), token(""), delimiter(DEFAULT_DELIMITER)
{
    currPos = buffer.begin();
}

/**
 * @details
 */
XLTokenizer::XLTokenizer(const std::string& str, const std::string& delimiter) : buffer(str), token(""), delimiter(delimiter)
{
    currPos = buffer.begin();
}

/**
 * @details
 */
void XLTokenizer::set(const std::string& str, const std::string& delimiter)
{
    this->buffer = str;
    this->delimiter = delimiter;
    this->currPos = buffer.begin();
}

/**
 * @details
 */
void XLTokenizer::setString(const std::string& str)
{
    this->buffer = str;
    this->currPos = buffer.begin();
}

/**
 * @details
 */
void XLTokenizer::setDelimiter(const std::string& delimiter)
{
    this->delimiter = delimiter;
    this->currPos = buffer.begin();
}

/**
 * @details
 */
std::string XLTokenizer::next()
{
    if(buffer.empty()) return "";           // skip if buffer is empty

    token.clear();                              // reset token string

    this->skipDelimiter();                      // skip leading delimiters

    // append each char to token string until it meets delimiter
    while(currPos != buffer.end() && !isDelimiter(*currPos))
    {
        token += *currPos;
        ++currPos;
    }
    return token;
}

/**
 * @details
 */
void XLTokenizer::skipDelimiter()
{
    while(currPos != buffer.end() && isDelimiter(*currPos))
        ++currPos;
}

/**
 * @details
 */
bool XLTokenizer::isDelimiter(char c)
{
    return (delimiter.find(c) != std::string::npos);
}

/**
 * @details
 */
std::vector<std::string> XLTokenizer::split()
{
    std::vector<std::string> tokens;
    std::string token;
    while((token = this->next()) != "")
    {
        tokens.push_back(token);
    }

    return tokens;
}
