//
// Created by Kenneth Balslev on 14/08/2020.
//

#include <pybind11/pybind11.h>
#include <pybind11/stl.h>
#include <XLSheet.hpp>
#include <XLCellReference.hpp>
#include <XLCellRange.hpp>
#include <XLCell.hpp>

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