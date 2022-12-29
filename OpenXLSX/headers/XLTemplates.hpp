//
// Created by Akira215 on 28/12/2022.
//

#ifndef OPENXLSX_XLTEMPLATES_HPP
#define OPENXLSX_XLTEMPLATES_HPP

#include <string> 

namespace OpenXLSX
{
    const std::string emptyWorksheetTo {
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
                " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
                " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\""
                " xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">"
                "<dimension ref=\"A1\"/>"
                "<sheetViews>"
                "<sheetView workbookViewId=\"0\"/>"
                "</sheetViews>"
                "<sheetFormatPr baseColWidth=\"10\" defaultRowHeight=\"16\" x14ac:dyDescent=\"0.2\"/>"
                "<sheetData/>"
                "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
                "</worksheet>"
                };
    
    const std::string emptyTable {
                "(<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
                "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
                " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"xr xr3\""
                " xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\""
                " xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\">"
                "<tableColumns></tableColumns><tableStyleInfo name=\"TableStyleLight1\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\"/></table>"
                };

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLTEMPLATES_HPP
