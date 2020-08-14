//
// Created by Kenneth Balslev on 14/08/2020.
//

#include <pybind11/pybind11.h>
#include <pybind11/stl.h>
#include <XLCellValue.hpp>

namespace py = pybind11;
using namespace OpenXLSX;

void init_XLValueType(py::module &m) {
    py::enum_<XLValueType>(m, "XLValueType")
        .value("Empty", XLValueType::Empty)
        .value("Boolean", XLValueType::Boolean)
        .value("Integer", XLValueType::Integer)
        .value("Float", XLValueType::Float)
        .value("Error", XLValueType::Error)
        .value("String", XLValueType::String)
        ;
}

void init_XLCellValue(py::module &m) {
    py::class_<XLCellValue>(m, "XLCellValue")
        .def("setInteger", &XLCellValue::set<int64_t>)
        .def("setFloat", &XLCellValue::set<long double>)
        .def("setBoolean", &XLCellValue::set<bool>)
        //.def("setString", py::overload_cast<const std::string&>(&XLCellValue::set))
        .def("getInteger", &XLCellValue::get<int64_t>)
        .def("getFloat", &XLCellValue::get<long double>)
        .def("getBoolean", &XLCellValue::get<bool>)
        .def("getString", &XLCellValue::get<const std::string &>)
        .def("asString", &XLCellValue::asString)
        .def("clear", &XLCellValue::clear)
        .def("valueType", &XLCellValue::valueType)
        ;
}

