/*

   ____                               ____      ___ ____       ____  ____      ___
  6MMMMb                              `MM(      )M' `MM'      6MMMMb\`MM(      )M'
 8P    Y8                              `MM.     d'   MM      6M'    ` `MM.     d'
6M      Mb __ ____     ____  ___  __    `MM.   d'    MM      MM        `MM.   d'
MM      MM `M6MMMMb   6MMMMb `MM 6MMb    `MM. d'     MM      YM.        `MM. d'
MM      MM  MM'  `Mb 6M'  `Mb MMM9 `Mb    `MMd       MM       YMMMMb     `MMd
MM      MM  MM    MM MM    MM MM'   MM     dMM.      MM           `Mb     dMM.
MM      MM  MM    MM MMMMMMMM MM    MM    d'`MM.     MM            MM    d'`MM.
YM      M9  MM    MM MM       MM    MM   d'  `MM.    MM            MM   d'  `MM.
 8b    d8   MM.  ,M9 YM    d9 MM    MM  d'    `MM.   MM    / L    ,M9  d'    `MM.
  YMMMM9    MMYMMM9   YMMMM9 _MM_  _MM_M(_    _)MM_ _MMMMMMM MYMMMM9 _M(_    _)MM_
            MM
            MM
           _MM_

  Copyright (c) 2020, Kenneth Troldal Balslev

  All rights reserved.

  Redistribution and use in source and binary forms, with or without
  modification, are permitted provided that the following conditions are met:
  - Redistributions of source code must retain the above copyright
    notice, this list of conditions and the following disclaimer.
  - Redistributions in binary form must reproduce the above copyright
    notice, this list of conditions and the following disclaimer in the
    documentation and/or other materials provided with the distribution.
  - Neither the name of the author nor the
    names of any contributors may be used to endorse or promote products
    derived from this software without specific prior written permission.

  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
  ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
  WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
  DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY
  DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
  (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
  LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
  (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
  SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

 */

#include <pybind11/pybind11.h>
#include <pybind11/stl.h>
#include <XLSheet.hpp>
#include <XLCellReference.hpp>
#include <XLCellRange.hpp>

namespace py = pybind11;
using namespace OpenXLSX;

void init_XLSheetState(py::module &m) {
    py::enum_<XLSheetState>(m, "XLSheetState")
        .value("Visible", XLSheetState::Visible)
        .value("Hidden", XLSheetState::Hidden)
        .value("VeryHidden", XLSheetState::VeryHidden);
}

void init_XLWorksheet(py::module &m) {
    py::class_<XLWorksheet>(m, "XLWorksheet")
        .def("visibility", &XLWorksheet::visibility, "Get the visibility state for the sheet.")
        .def("setVisibility", &XLWorksheet::setVisibility, "Set the visibility state for the sheet.", py::arg("visibility"))
        .def("color", &XLWorksheet::color, "Get the color of the sheet.")
        .def("setColor", &XLWorksheet::setColor, "Set the color of the sheet.", py::arg("color"))
        .def("index", &XLWorksheet::index, "Get the index of the sheet.")
        .def("setIndex", &XLWorksheet::setIndex, "Set the index of the sheet.", py::arg("index"))
        .def("name", &XLWorksheet::name, "Get the name of the sheet.")
        .def("setName", &XLWorksheet::setName, "Set the name of the sheet.", py::arg("name"))
        .def("isSelected", &XLWorksheet::isSelected, "Determine if the sheet is currently selected in the workbook.")
        .def("setSelected", &XLWorksheet::setSelected, "Set the 'selected' state for the sheet.", py::arg("selected"))
        .def("clone", &XLWorksheet::clone, "Clone the sheet.", py::arg("newName"))
        .def("cell", py::overload_cast<const std::string&>(&XLWorksheet::cell), "Get the cell with the given address.", py::arg("address"))
        .def("cell",
             py::overload_cast<const XLCellReference&>(&XLWorksheet::cell),
             "Get the cell with the given reference.",
             py::arg("reference"))
        .def("cell",
             py::overload_cast<uint32_t, uint16_t>(&XLWorksheet::cell),
             "Get the cell with the given coordinates",
             py::arg("row"),
             py::arg("column"))
        .def("range", py::overload_cast<>(&XLWorksheet::range), "Get a range object for the range of cells currently in use.")
        .def("range",
             py::overload_cast<const XLCellReference&, const XLCellReference&>(&XLWorksheet::range),
             "Get a range object with the given coordinates.",
             py::arg("topLeft"),
             py::arg("bottomRight"))
        .def("row", &XLWorksheet::row, "Get the row object for the given row number..", py::arg("rowNumber"))
        .def("column", &XLWorksheet::column, "Get the column object for the given column number.", py::arg("columnNumber"))
        .def("columnCount", &XLWorksheet::columnCount, "Get the number of columns in the worksheet.")
        .def("rowCount", &XLWorksheet::rowCount, "Get the number of rows in the worksheet.");
}

void init_XLChartsheet(py::module &m) {
    py::class_<XLChartsheet>(m, "XLChartsheet")
        .def("visibility", &XLChartsheet::visibility, "Get the visibility state for the sheet.")
        .def("setVisibility", &XLChartsheet::setVisibility, "Set the visibility state for the sheet.", py::arg("visibility"))
        .def("color", &XLChartsheet::color, "Get the color of the sheet.")
        .def("setColor", &XLChartsheet::setColor, "Set the color of the sheet.", py::arg("color"))
        .def("index", &XLChartsheet::index, "Get the index of the sheet.")
        .def("setIndex", &XLChartsheet::setIndex, "Set the index of the sheet.", py::arg("index"))
        .def("name", &XLChartsheet::name, "Get the name of the sheet.")
        .def("setName", &XLChartsheet::setName, "Set the name of the sheet.", py::arg("name"))
        .def("isSelected", &XLChartsheet::isSelected, "Determine if the sheet is currently selected in the workbook.")
        .def("setSelected", &XLChartsheet::setSelected, "Set the 'selected' state for the sheet.", py::arg("selected"))
        .def("clone", &XLChartsheet::clone, "Clone the sheet.", py::arg("newName"));
}

void init_XLSheet(py::module &m) {
    py::class_<XLSheet>(m, "XLSheet")
        .def("visibility", &XLSheet::visibility, "Get the visibility state for the sheet.")
        .def("setVisibility", &XLSheet::setVisibility, "Set the visibility state for the sheet.", py::arg("visibility"))
        .def("color", &XLSheet::color, "Get the color of the sheet.")
        .def("setColor", &XLSheet::setColor, "Set the color of the sheet.", py::arg("color"))
        .def("index", &XLSheet::index, "Get the index of the sheet.")
        .def("setIndex", &XLSheet::setIndex, "Set the index of the sheet.", py::arg("index"))
        .def("name", &XLSheet::name, "Get the name of the sheet.")
        .def("setName", &XLSheet::setName, "Set the name of the sheet.", py::arg("name"))
        .def("setSelected", &XLSheet::setSelected, "Set the 'selected' state for the sheet.", py::arg("selected"))
        .def("clone", &XLSheet::clone, "Clone the sheet.", py::arg("newName"))
        .def_property_readonly(
            "type",
            [](const XLSheet& sheet) {
                if (sheet.isType<XLWorksheet>())
                    return XLSheetType::Worksheet;
                else
                    return XLSheetType::Chartsheet;
            },
            "Get the underlying sheet type (worksheet or chartsheet).")
        .def("worksheet", &XLSheet::get<XLWorksheet>, "Get the underlying worksheet object.")
        .def("chartsheet", &XLSheet::get<XLChartsheet>, "Get the underlying chartsheet object.");
}