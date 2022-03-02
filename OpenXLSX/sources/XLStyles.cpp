// ===== External Includes ===== //
#include <algorithm>
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLCell.hpp"
#include "XLDocument.hpp"
#include "XLStyles.hpp"
#include "XLXmlData.hpp"

using namespace OpenXLSX;




#define AMOUNT_OF_PREDEFINED_STYLE_TYPES 50


// ================================================================================
// Constant array from xlsx_drone
// ================================================================================
static enum CellCategory xlsx_predefined_style_types[AMOUNT_OF_PREDEFINED_STYLE_TYPES] = {
  XLSX_UNKNOWN, // 0
  XLSX_NUMBER, // 1
  XLSX_NUMBER, // 2
  XLSX_NUMBER, // 3
  XLSX_NUMBER, // 4
  XLSX_NUMBER, // 5
  XLSX_NUMBER, // 6
  XLSX_NUMBER, // 7
  XLSX_NUMBER, // 8
  XLSX_NUMBER, // 9
  XLSX_NUMBER, // 10
  XLSX_NUMBER, // 11
  XLSX_NUMBER, // 12
  XLSX_NUMBER, // 13
  XLSX_DATE, // 14
  XLSX_DATE, // 15
  XLSX_DATE, // 16
  XLSX_DATE, // 17
  XLSX_DATE_TIME, // 18
  XLSX_DATE_TIME, // 19
  XLSX_TIME, // 20
  XLSX_TIME, // 21
  XLSX_DATE_TIME, // 22
  XLSX_UNKNOWN, // 23
  XLSX_UNKNOWN, // 24
  XLSX_UNKNOWN, // 25
  XLSX_UNKNOWN, // 26
  XLSX_UNKNOWN, // 27
  XLSX_UNKNOWN, // 28
  XLSX_UNKNOWN, // 29
  XLSX_UNKNOWN, // 30
  XLSX_UNKNOWN, // 31
  XLSX_UNKNOWN, // 32
  XLSX_UNKNOWN, // 33
  XLSX_UNKNOWN, // 34
  XLSX_UNKNOWN, // 35
  XLSX_UNKNOWN, // 36
  XLSX_NUMBER, // 37
  XLSX_NUMBER, // 38
  XLSX_NUMBER, // 39
  XLSX_NUMBER, // 40
  XLSX_UNKNOWN, // 41
  XLSX_UNKNOWN, // 42
  XLSX_UNKNOWN, // 43
  XLSX_UNKNOWN, // 44
  XLSX_TIME, // 45
  XLSX_TIME, // 46
  XLSX_TIME, // 47
  XLSX_NUMBER, // 48
  XLSX_TEXT // 49
};

static const char *xlsx_predefined_styles_format_code[AMOUNT_OF_PREDEFINED_STYLE_TYPES] = {
    NULL, // 0
    "0", // 1
    "0.00", // 2
    "#,##0", // 3
    "#,##0.00", // 4
    "$#,##0_);($#,##0)", // 5
    "$#,##0_);[Red]($#,##0)", // 6
    "$#,##0.00_);($#,##0.00)", // 7
    "$#,##0.00_);[Red]($#,##0.00)", // 8
    "0%", // 9
    "0.00%", // 10
    "0.00E+00", // 11
    "# ?/?", // 12
    "# ?\?/??", // 13
    "d/m/yyyy", // 14
    "d-mmm-yy", // 15
    "d-mmm", // 16
    "mmm-yy", // 17
    "h:mm AM/PM", // 18
    "h:mm:ss AM/PM", // 19
    "h:mm", // 20
    "h:mm:ss", // 21
    "m/d/yyyy h:mm", // 22
    NULL, // 23
    NULL, // 24
    NULL, // 25
    NULL, // 26
    NULL, // 27
    NULL, // 28
    NULL, // 29
    NULL, // 30
    NULL, // 31
    NULL, // 32
    NULL, // 33
    NULL, // 34
    NULL, // 35
    NULL, // 36
    "#,##0_);(#,##0)", // 37
    "#,##0_);[Red](#,##0)", // 38
    "#,##0.00_);(#,##0.00)", // 39
    "#,##0.00_);[Red](#,##0.00)", // 40
    NULL, // 41
    NULL, // 42
    NULL, // 43
    NULL, // 44
    "mm:ss", // 45
    "[h]:mm:ss", // 46
    "mm:ss.0", // 47
    "##0.0E+0", // 48
    "@" // 49
};

/*
* summary:
*   Analyze an explicit char of the *format_code* and tells if it's part of some specific formatting.
* params:
*   - format_code: format_code withdrawn from a numFmt node in the styles XML.
*   - current_analyzed_index: the index pointing to a specific char of *format_code*.
*/
static enum Formatter get_formatter(const char *format_code, int current_analyzed_index) {
  switch(format_code[current_analyzed_index]) {
    case 'm': case 'h': case 's': case 'y': case 'd': {
      // "[Red]" case
      if(format_code[current_analyzed_index] == 'd' && current_analyzed_index >= 3 &&
        format_code[current_analyzed_index - 3] == '[' && format_code[current_analyzed_index - 2] == 'R' &&
        format_code[current_analyzed_index - 1] == 'e' && format_code[current_analyzed_index + 1] == ']') {
        return XLSX_FORMATTER_UNKNOWN;
      }

      // so far it isn't escaped
      int char_is_escaped = false;
      // check if it's escaped
      if(current_analyzed_index > 0) {

        // being preceded by '\'
        if(format_code[current_analyzed_index - 1] == '\\') {
          // check if the '\' is actually being escaped
          if(current_analyzed_index > 1) {

            // so far it is escaped
            char_is_escaped = true;
            int i = current_analyzed_index - 2;

            // look backwards until find a non '\' char or the start of the string
            while(true) {
              if(format_code[i] == '\\') {
                if(char_is_escaped) {
                  char_is_escaped = false;
                } else {
                  char_is_escaped = true;
                }
              } else {
                break;
              }

              if(--i < 0) {
                break;
              }
            }
          } else {
            char_is_escaped = true;
          }
        } else {

          // being surrounded by '"'
          int quotes_open = false;
          int i = 0;

          while(true) {
            switch(format_code[i]) {
              case '\0':
                break;
              case '"':
                if(!quotes_open) {
                  quotes_open = true;
                } else {
                  // see if our currently analyzed index is in-between
                  if(current_analyzed_index < i) {
                    char_is_escaped = true;
                    break;
                  }
                  quotes_open = false;
                }
            }

            if((++i == current_analyzed_index) && (!quotes_open)) {
              break;
            }
          }
        }

      }

      if(char_is_escaped) {
        return XLSX_FORMATTER_UNKNOWN;
      } else {
        switch(format_code[current_analyzed_index]) {
          case 'm':
            return XLSX_FORMATTER_AMBIGUOUS_M;
          case 'h': case 's':
            return XLSX_FORMATTER_TIME;
          default:
            // 'y' or 'd'
            return XLSX_FORMATTER_DATE;
        }
      }

    } default:
      return XLSX_FORMATTER_UNKNOWN;
  }
}

/*
* summary:
*   Parses a *format_code*, in search for clues, that takes the program to understand which is the type the
*   *format_code* is talking about. Don't pass *format_code* that formats raw text, this function expects to work
*   with numbers.
* params:
*   - format_code: format_code withdrawn from a numFmt node in the styles XML.
*   - format_code_length: for speeding purpose.
* returns:
*   One of:
*     - XLSX_NUMBER
*     - XLSX_DATE
*     - XLSX_TIME
*     - XLSX_DATE_TIME
*/
static enum CellCategory get_related_category(const char *format_code, int format_code_length) {
  // *m_found* means an 'm' char found, it's ambiguous between date and time
  int current_analyzed_index, is_date = 0, is_time = 0, m_found = 0;
  for(current_analyzed_index = 0; current_analyzed_index < format_code_length; ++current_analyzed_index) {

    // note that this could also return XLSX_FORMATTER_UNKNOWN
    switch(get_formatter(format_code, current_analyzed_index)) {
      case XLSX_FORMATTER_AMBIGUOUS_M:
        m_found = true;
        break;
      case XLSX_FORMATTER_TIME:
        is_time = true;
        break;
      case XLSX_FORMATTER_DATE:
        is_date = true;
        break;
      default: {}
    }

    if(is_date && is_time)
      // there's no need to keep researching
      return XLSX_DATE_TIME;
  }

  // inspect the result of the research
  if(is_time) {
    return XLSX_TIME;
  } else if(is_date || m_found) {
    return XLSX_DATE;
  } else {
    return XLSX_NUMBER;
  }
}

static enum Formatter get_formatter_null(const char* format_code)
{
    return (format_code==NULL) ? XLSX_FORMATTER_UNKNOWN : get_formatter(format_code,static_cast<int>(strlen(format_code)));
}

// ================================================================================
// Helpers
// ================================================================================
template<typename Out>
inline void splitA(const std::string& s, char delim, Out result)
{
    std::istringstream iss(s);
    std::string        item;
    while (std::getline(iss, item, delim)) {
        *result++ = item;
    }
}

inline void splitA(const std::string& s, char delim, std::vector<std::string>& elems)
{
    splitA(s, delim, std::back_inserter(elems));
}

static pugi::xml_node findFontNode(const XMLDocument& xmlDoc, int fontId)
{
    //xmlDoc.find_child() ? how to use
    pugi::xml_node             node;
    constexpr std::string_view nodeFontsName { "fonts" };
    constexpr std::string_view nodeFontName { "font" };
    for (const auto& nodeItem : xmlDoc.document_element().children()) {
        if (nodeFontsName == nodeItem.name()) {
            for (auto& xfNode : nodeItem.children()) {
                if (nodeFontName == xfNode.name()) {
                    if (fontId == 0) {
                        return xfNode;
                    }
                    else {
                        int idx = 0;
                        while (idx < fontId) {
                            xfNode = xfNode.next_sibling();
                            idx++;
                        }
                        return xfNode;
                    }
                }
            }
        }
    }
    return node;
}

// ================================================================================
// XLStyles Class
// ================================================================================

XLStyles::XLStyles(XLXmlData* xmlData) : XLXmlFile(xmlData)
{
    init(xmlData);
}

XLStyles::~XLStyles() = default;

OpenXLSX::XLStyle OpenXLSX::XLStyles::style(const XLCell& cell) const
{
    XLStyle res(this->parentDoc());
    if (!cell) return res;
    if (const auto* val = cell.m_cellNode->attribute("s").value(); val != nullptr) {
        if (const size_t pos = std::atoi(val); m_VecStyle.size() > pos) {
            res = m_VecStyle.at(pos);
        }
    }
    return res;
}

std::string OpenXLSX::XLStyles::formatString(int numFmtId) const
{
    constexpr const auto* snumFmts    = "numFmts";
    constexpr const auto* snumFmtId   = "numFmtId";
    constexpr const auto* sformatCode = "formatCode";

    return xmlDocument()
        .document_element()
        .child(snumFmts)
        .find_child_by_attribute(snumFmtId, std::to_string(numFmtId).c_str())
        .attribute(sformatCode)
        .value();
}


enum Formatter OpenXLSX::XLStyles::formatter(int numFmtId) const
{
    if (numFmtId < AMOUNT_OF_PREDEFINED_STYLE_TYPES) {
        // predefined style
        return get_formatter_null(xlsx_predefined_styles_format_code[numFmtId]);
    }
    else {
        std::string str(formatString(numFmtId));
        return get_formatter(str.c_str(), static_cast<int>(str.length()));
    }
}

enum CellCategory OpenXLSX::XLStyles::cellCategory(int numFmtId) const
{
    if (numFmtId < AMOUNT_OF_PREDEFINED_STYLE_TYPES) {
        // predefined style
        return xlsx_predefined_style_types[numFmtId];
    }
    else {
        std::string str(formatString(numFmtId));
        return get_related_category(str.c_str(), static_cast<int>(str.length()));
    }
}

void OpenXLSX::XLStyles::init(const XLXmlData* stylesData)
{
    // cache cellXfs
    if (stylesData) {
        constexpr const auto* numFmtId = "numFmtId";
        constexpr const auto* fontId   = "fontId";
        constexpr const auto* fillId   = "fillId";
        constexpr const auto* borderId = "borderId";
        constexpr const auto* xfId     = "xfId";

        constexpr std::string_view nodeNameCellXfs { "cellXfs" };
        constexpr std::string_view nodeNameXf { "xf" };

        // TODO: check applyNumberFormat, reserve size from cellXfs count;
        for (const auto& node : stylesData->getXmlDocument()->document_element().children()) {
            if (nodeNameCellXfs == node.name()) {
                for (const auto& xfNode : node.children()) {
                    if (nodeNameXf == xfNode.name()) {
                        XLStyle item(this->parentDoc());
                        if (const auto* val = xfNode.attribute(numFmtId).value(); val != nullptr) item.m_numFmtId = std::stoi(val);
                        if (const auto* val = xfNode.attribute(fontId).value(); val != nullptr) item.m_fontId = std::stoi(val);
                        if (const auto* val = xfNode.attribute(fillId).value(); val != nullptr) item.m_fillId = std::stoi(val);
                        if (const auto* val = xfNode.attribute(borderId).value(); val != nullptr) item.m_borderId = std::stoi(val);
                        if (const auto* val = xfNode.attribute(xfId).value(); val != nullptr) item.m_xfId = std::stoi(val);
                        m_VecStyle.emplace_back(item);
                    }
                }
                break;
            }
            // TODO: Fonts
        }
    }
}

// ================================================================================
// XLStyle Class
// ================================================================================

OpenXLSX::XLStyle::XLStyle(const XLDocument& doc) : m_doc(doc) {}

OpenXLSX::XLNumberFormat OpenXLSX::XLStyle::numberFormat() const
{
    XLNumberFormat fmt(*this);
    return fmt;
}

std::string OpenXLSX::XLStyle::formatString() const
{
    return m_doc.get().styles().formatString(m_numFmtId);
}

int OpenXLSX::XLStyle::numFmtId() const
{
    return m_numFmtId;
}

OpenXLSX::XLFont OpenXLSX::XLStyle::font() const
{
    return XLFont(*this, findFontNode(m_doc.get().styles().xmlDocument(), m_fontId));
}

// ================================================================================
// XLFont Class
// ================================================================================
// XLQuery?
OpenXLSX::XLFont::XLFont(const XLStyle& style, const XMLNode& node) : 
    m_style(style), m_node(new XMLNode(node)) {}

OpenXLSX::XLFont::~XLFont() {}

std::string OpenXLSX::XLFont::name() const
{
    std::string _name;
    if (isValid()) {
        _name = m_node->child("name").attribute("val").value();
    }
    return _name;
}

double OpenXLSX::XLFont::size() const
{
    double result = 0.0;
    if (isValid()) {
        result = std::stod(m_node->child("sz").attribute("val").value());
    }
    return result;
}

OpenXLSX::XLColor OpenXLSX::XLFont::color() const
{
    OpenXLSX::XLColor clr;
    if (isValid()) {
        const std::string val = m_node->child("color").attribute("rgb").value();
        if (val.size() != 0) {
            clr.set(val);
        }
    }
    return clr;
}

int OpenXLSX::XLFont::colorIndex() const
{
    constexpr int xlColorIndexNone = -4142;
    if (isValid()) {
        const std::string val = m_node->child("color").attribute("theme").value();
        if (val.size() != 0) {
            return std::atoi(val.c_str());
        }
    }
    return xlColorIndexNone;
}


int OpenXLSX::XLFont::underline() const {
    //https://docs.microsoft.com/en-us/office/vba/api/excel.xlunderlinestyle
    if (isValid()) {
        constexpr std::string_view nodeName { "u" };
        constexpr std::string_view attDouble { "double" };
        for (const auto& node : m_node->children()) {

            if (nodeName == node.name()) {
                if (auto attribute = node.attribute("val"); !attribute.empty() && attribute.value() == attDouble) {
                    return -4119;
                }
                else {
                    return 2;
                }

            }

        }
    }
    return -4142;
}

bool OpenXLSX::XLFont::strikethrough() const {
    if (isValid()) {
        constexpr std::string_view nodeName { "strike" };
        for (const auto& node : m_node->children()) {
            if (nodeName == node.name()) {
                return true;
            }
        }
    }
    return false;
}


bool OpenXLSX::XLFont::bold() const {
    if (isValid()) {
        constexpr std::string_view nodeName { "b" };
        for (const auto& node : m_node->children()) {
            if (nodeName == node.name()) {
                return true;
            }
        }
    }
    return false;
}


bool OpenXLSX::XLFont::italic() const{
    if (isValid()) {
        constexpr std::string_view nodeName { "i" };
        for (const auto& node : m_node->children()) {
            if (nodeName == node.name()) {
                return true;
            }
        }
    }
    return false;
}

bool OpenXLSX::XLFont::isValid() const 
{
    constexpr std::string_view nodeFontName { "font" };
    if (m_style.get().m_fontId != -1 && m_node != nullptr && m_node->type() != pugi::xml_node_type::node_null && m_node->name() == nodeFontName){
        return true;
    }
    return false;
}

// ================================================================================
// XLNumberFormat Class
// ================================================================================
OpenXLSX::XLNumberFormat::XLNumberFormat(const XLStyle& style) : m_style { style } {}

OpenXLSX::XLNumberFormat::XLNumberType OpenXLSX::XLNumberFormat::type()
{
    return tryFindType();
}

std::string OpenXLSX::XLNumberFormat::currencySumbol() const
{
    return m_currencySumbol;
}

// TODO: it may be possible to get precision and forrmatting in a single pass.
OpenXLSX::XLNumberFormat::XLNumberType OpenXLSX::XLNumberFormat::tryFindType()
{
    const auto _type = tryBuiltinType();
    if (_type != XLNumberType::kUnkown) return _type;

    const std::string fmt = m_style.get().formatString();
    if (fmt.size() == 0) return XLNumberType::kUnkown;

    std::vector<std::string> bracketTextArray;
    std::lconv*              lc = std::localeconv();
    std::string              bracketText;
    bool                     isInBracket = false;

    auto foundType = XLNumberType::kUnkown;

    if (fmt.find("AM/PM") != std::string::npos) return XLNumberType::kDate;

    for (size_t idx = 0; idx < fmt.size(); idx++) {
        if (fmt[idx] == '[') {
            isInBracket = true;
        }
        else if (fmt[idx] == ']') {
            isInBracket = false;
            bracketTextArray.push_back(bracketText);
            bracketText.clear();
        }
        else if (isInBracket) {
            bracketText += fmt[idx];
        }
        else {
            const char chr = fmt[idx];
            if (chr == '$') {
                m_currencySumbol = chr;
                return XLNumberType::kCurrency;
            }
            else if (std::string lc_curr = lc->currency_symbol; lc_curr.find(chr) != std::string::npos) {
                m_currencySumbol = lc->currency_symbol;
                return XLNumberType::kCurrency;
            }
            else if (chr == '%') {
                return XLNumberType::kPercent;
            }
            else if (chr == 'y' || chr == 'M' || chr == 'm' || chr == 'd') {
                return XLNumberType::kDate;
            }
            else if (chr == 'h' || chr == 's' || chr == ':') {
                return XLNumberType::kDate;
            }
            else if (chr == '0' || chr == '#' || chr == '.') {
                foundType = XLNumberType::kUnkown;
            }
            // use lc->int_curr_symbol ? how to get en_US?
            //(iCompare(m_slocale, L"ja_jp") || iCompare(m_slocale, L"zh_tw")) && wchr == 'e' || wchr == 'g' || wchr == 'r')
        }
    }
    for (const auto& item : bracketTextArray) {
        if (item.find("F800") != std::string::npos) {
            return XLNumberType::kDate;
        }
        else if (item.find("F400") != std::string::npos) {
            return XLNumberType::kDate;
        }
        else {
            if (size_t pos = item.find('-'); pos != std::wstring::npos) {
                std::vector<std::string> parts;
                splitA(item, '-', parts);
                if (parts.size() == 2) {
                    if (parts[0].size() > 1) {
                        m_currencySumbol = parts[0].substr(1, parts[0].size() - 1);
                        m_fmtLocal       = parts[1];
                        return XLNumberType::kCurrency;
                    }
                }
            }
            else {
                if (item[0] == '$') {
                    m_currencySumbol = item.substr(1, item.size() - 1);
                    return XLNumberType::kCurrency;
                }
            }
        }
    }
    return foundType;
}

OpenXLSX::XLNumberFormat::XLNumberType OpenXLSX::XLNumberFormat::tryBuiltinType()
{
    std::setlocale(LC_MONETARY, "");
    std::lconv* lc = std::localeconv();
    switch (m_style.get().numFmtId()) {
        case 0:
        case 1:
        case 3:
            return XLNumberType::kUnkown;
        case 2:
        case 4:
            return XLNumberType::kUnkown;
        case 5:
        case 6:
        case 7:
        case 8:
            m_currencySumbol = lc->currency_symbol;
            return XLNumberType::kCurrency;
        case 9:
        case 10:
            return XLNumberType::kPercent;
        case 11:
        case 12:
        case 13:
            return XLNumberType::kUnkown;
        case 14:
        case 15:
        case 16:
        case 17:
        case 18:
        case 19:
        case 20:
        case 21:
        case 22:
            return XLNumberType::kDate;
        case 37:
        case 38:
        case 39:
        case 40:    // accounting
            return XLNumberType::kUnkown;
        case 45:
        case 46:
        case 47:
        case 48:
            return XLNumberType::kDate;
    }
    return XLNumberType::kUnkown;
}