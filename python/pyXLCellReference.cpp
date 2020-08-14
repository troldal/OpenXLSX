//
// Created by Kenneth Balslev on 14/08/2020.
//

#include <pybind11/pybind11.h>
#include <XLCellReference.hpp>

namespace py = pybind11;
using namespace OpenXLSX;

void init_XLCellReference(py::module &m) {
    py::class_<XLCellReference>(m, "XLCellReference")
        .def(py::init<>())
        .def(py::init<const std::string &>())
        .def(py::init<uint32_t, uint16_t>())
        .def(py::init<uint32_t, const std::string &>())
        .def("row", &XLCellReference::row)
        .def("setRow", &XLCellReference::setRow)
        .def("column", &XLCellReference::column)
        .def("setColumn", &XLCellReference::setColumn)
        .def("setRowAndColumn", &XLCellReference::setRowAndColumn)
        .def("address", &XLCellReference::address)
        .def("setAddress", &XLCellReference::setAddress);
}

