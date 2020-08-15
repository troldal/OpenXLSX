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

  Copyright (c) 2018, Kenneth Troldal Balslev

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
        .def("visibility", &XLWorksheet::visibility)
        .def("setVisibility", &XLWorksheet::setVisibility)
        .def("color", &XLWorksheet::color)
        .def("setColor", &XLWorksheet::setColor)
        .def("index", &XLWorksheet::setIndex)
        .def("setIndex", &XLWorksheet::setIndex)
        .def("name", &XLWorksheet::name)
        .def("setName", &XLWorksheet::setName)
        .def("isSelected", &XLWorksheet::isSelected)
        .def("setSelected", &XLWorksheet::setSelected)
        .def("clone", &XLWorksheet::clone)
        .def("cell", py::overload_cast<const std::string&>(&XLWorksheet::cell))
        .def("cell", py::overload_cast<const XLCellReference &>(&XLWorksheet::cell))
        .def("cell", py::overload_cast<uint32_t, uint16_t>(&XLWorksheet::cell))
        .def("range", py::overload_cast<>(&XLWorksheet::range))
        .def("range", py::overload_cast<const XLCellReference&, const XLCellReference&>(&XLWorksheet::range))
        .def("row", &XLWorksheet::row)
        .def("column", &XLWorksheet::column)
        .def("lastCell", &XLWorksheet::lastCell)
        .def("columnCount", &XLWorksheet::columnCount)
        .def("rowCount", &XLWorksheet::rowCount)
        ;
}

void init_XLChartsheet(py::module &m) {
    py::class_<XLChartsheet>(m, "XLChartsheet")
        .def("visibility", &XLChartsheet::visibility)
        .def("setVisibility", &XLChartsheet::setVisibility)
        .def("color", &XLChartsheet::color)
        .def("setColor", &XLChartsheet::setColor)
        .def("index", &XLChartsheet::setIndex)
        .def("setIndex", &XLChartsheet::setIndex)
        .def("name", &XLChartsheet::name)
        .def("setName", &XLChartsheet::setName)
        .def("isSelected", &XLChartsheet::isSelected)
        .def("setSelected", &XLChartsheet::setSelected)
        .def("clone", &XLChartsheet::clone)
        ;
}

void init_XLSheet(py::module &m) {
    py::class_<XLSheet>(m, "XLSheet")
        .def("visibility", &XLSheet::visibility)
        .def("setVisibility", &XLSheet::setVisibility)
        .def("color", &XLSheet::color)
        .def("setColor", &XLSheet::setColor)
        .def("index", &XLSheet::setIndex)
        .def("setIndex", &XLSheet::setIndex)
        .def("name", &XLSheet::name)
        .def("setName", &XLSheet::setName)
        .def("setSelected", &XLSheet::setSelected)
        .def("isWorksheet", &XLSheet::isType<XLWorksheet>)
        .def("isChartsheet", &XLSheet::isType<XLChartsheet>)
        .def("clone", &XLSheet::clone)
        .def("worksheet", &XLSheet::get<XLWorksheet>)
        .def("chartsheet", &XLSheet::get<XLChartsheet>)
        ;
}