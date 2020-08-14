//
// Created by Kenneth Balslev on 14/08/2020.
//

#include <pybind11/pybind11.h>
#include <XLCell.hpp>

namespace py = pybind11;
using namespace OpenXLSX;

void init_XLCell(py::module &m) {
    py::class_<XLCell>(m, "XLCell")
        .def("value", &XLCell::value)
        .def("valueType", &XLCell::valueType)
        .def("cellReference", &XLCell::cellReference)
        .def("hasFormula", &XLCell::hasFormula)
        .def("formula", &XLCell::formula)
        .def("setFormula", &XLCell::setFormula)
        ;
}

