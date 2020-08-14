//
// Created by Kenneth Balslev on 14/08/2020.
//

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
        .def("sheet", py::overload_cast<uint16_t>(&XLWorkbook::sheet))
        .def("sheet", py::overload_cast<const std::string&>(&XLWorkbook::sheet))
        .def("worksheet", &XLWorkbook::worksheet)
        .def("chartsheet", &XLWorkbook::chartsheet)
        .def("deleteSheet", &XLWorkbook::deleteSheet)
        .def("addWorksheet", &XLWorkbook::addWorksheet)
        .def("cloneSheet", &XLWorkbook::cloneSheet)
        .def("setSheetIndex", &XLWorkbook::setSheetIndex)
        .def("typeOfSheet", py::overload_cast<unsigned int>(&XLWorkbook::typeOfSheet, py::const_))
        .def("typeOfSheet", py::overload_cast<const std::string&>(&XLWorkbook::typeOfSheet, py::const_))
        .def("sheetCount", &XLWorkbook::sheetCount)
        .def("worksheetCount", &XLWorkbook::worksheetCount)
        .def("chartsheetCount", &XLWorkbook::chartsheetCount)
        .def("sheetNames", &XLWorkbook::sheetNames)
        .def("worksheetNames", &XLWorkbook::worksheetNames)
        .def("chartsheetNames", &XLWorkbook::chartsheetNames)
        .def("sheetExists", &XLWorkbook::sheetExists)
        .def("worksheetExists", &XLWorkbook::worksheetExists)
        .def("chartsheetExists", &XLWorkbook::chartsheetExists)
        .def("deleteNamedRanges", &XLWorkbook::deleteNamedRanges);
}