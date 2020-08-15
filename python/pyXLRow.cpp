//
// Created by Kenneth Balslev on 14/08/2020.
//

#include <XLRow.hpp>
#include <pybind11/pybind11.h>

namespace py = pybind11;
using namespace OpenXLSX;

void init_XLRow(py::module& m)
{
    py::class_<XLRow>(m, "XLRow")
        .def("height", &XLRow::height)
        .def("setHeight", &XLRow::setHeight)
        .def("descent", &XLRow::descent)
        .def("setDescent", &XLRow::setDescent)
        .def("isHidden", &XLRow::isHidden)
        .def("setHidden", &XLRow::setHidden)
        .def("rowNumber", &XLRow::rowNumber)
        .def("cellCount", &XLRow::cellCount);
}
