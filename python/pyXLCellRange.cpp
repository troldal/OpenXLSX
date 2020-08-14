//
// Created by Kenneth Balslev on 14/08/2020.
//

#include <pybind11/pybind11.h>
#include <XLCellRange.hpp>

namespace py = pybind11;
using namespace OpenXLSX;

void init_XLCellRange(py::module &m) {
    py::class_<XLCellRange>(m, "XLCellRange")
        .def("numRows", &XLCellRange::numRows)
        .def("numColumns", &XLCellRange::numColumns)
        .def("begin", &XLCellRange::begin)
        .def("end", &XLCellRange::end)
        .def("clear", &XLCellRange::clear)
        ;
}

