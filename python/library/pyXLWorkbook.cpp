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
#include <XLWorkbook.hpp>
#include <XLSheet.hpp>

namespace py = pybind11;
using namespace OpenXLSX;

void init_XLSheetType(py::module &m) {
    py::enum_<XLSheetType>(m, "XLSheetType")
        .value("Worksheet", XLSheetType::Worksheet)
        .value("Chartsheet", XLSheetType::Chartsheet)
        .value("Dialogsheet", XLSheetType::Dialogsheet)
        .value("Macrosheet", XLSheetType::Macrosheet);
}

void init_XLWorkbook(py::module &m) {
    py::class_<XLWorkbook>(m, "XLWorkbook")
        .def("sheet", py::overload_cast<uint16_t>(&XLWorkbook::sheet), "Get the sheet at the given index.", py::arg("index"))
        .def("sheet", py::overload_cast<const std::string&>(&XLWorkbook::sheet), "Get the sheet with the given name.", py::arg("name"))
        .def("worksheet", &XLWorkbook::worksheet, "Get the worksheet with the given name.", py::arg("name"))
        .def("chartsheet", &XLWorkbook::chartsheet, "Get the chartsheet with the given name.", py::arg("name"))
        .def("deleteSheet", &XLWorkbook::deleteSheet, "Delete the sheet with the given name.", py::arg("name"))
        .def("addWorksheet", &XLWorkbook::addWorksheet, "Add a new worksheet with the given name.", py::arg("name"))
        .def("cloneSheet", &XLWorkbook::cloneSheet, "Clone the sheet with the given name.", py::arg("existingName"), py::arg("newName"))
        .def("setSheetIndex",
             &XLWorkbook::setSheetIndex,
             "Set the index of the sheet with the given name.",
             py::arg("name"),
             py::arg("index"))
        .def("indexOfSheet", &XLWorkbook::indexOfSheet, "Get the index of the sheet with the given name.", py::arg("name"))
        .def("typeOfSheet",
             py::overload_cast<unsigned int>(&XLWorkbook::typeOfSheet, py::const_),
             "Get the type of sheet with the given index.",
             py::arg("index"))
        .def("typeOfSheet",
             py::overload_cast<const std::string&>(&XLWorkbook::typeOfSheet, py::const_),
             "Get the type of sheet with the given name.",
             py::arg("name"))
        .def("sheetCount", &XLWorkbook::sheetCount, "Get the total sheet count.")
        .def("worksheetCount", &XLWorkbook::worksheetCount, "Get the worksheet count.")
        .def("chartsheetCount", &XLWorkbook::chartsheetCount, "Get the chartsheet count.")
        .def("sheetNames", &XLWorkbook::sheetNames, "Get all the sheet names.")
        .def("worksheetNames", &XLWorkbook::worksheetNames, "Get the worksheet names.")
        .def("chartsheetNames", &XLWorkbook::chartsheetNames, "Get the chartsheet names.")
        .def("sheetExists", &XLWorkbook::sheetExists, "Determine if a sheet with the given name exists in the workbook.", py::arg("name"))
        .def("worksheetExists",
             &XLWorkbook::worksheetExists,
             "Determine if a worksheet with the given name exists in the workbook.",
             py::arg("name"))
        .def("chartsheetExists",
             &XLWorkbook::chartsheetExists,
             "Determine if a chartsheet with the given name exists in the workbook.",
             py::arg("name"))
        .def("deleteNamedRanges", &XLWorkbook::deleteNamedRanges, "Delete all named ranges in the workbook.");
}