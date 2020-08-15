//
// Created by Kenneth Balslev on 14/08/2020.
//

#include <XLColumn.hpp>
#include <pybind11/pybind11.h>

namespace py = pybind11;
using namespace OpenXLSX;

void init_XLColumn(py::module& m)
{
    py::class_<XLColumn>(m, "XLColumn")
        .def("width", &XLColumn::width)
        .def("setWidth", &XLColumn::setWidth)
        .def("isHidden", &XLColumn::isHidden)
        .def("setHidden", &XLColumn::setHidden);
}
