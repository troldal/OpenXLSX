//
// Created by Kenneth Balslev on 14/08/2020.
//

#include <XLColor.hpp>
#include <pybind11/pybind11.h>

namespace py = pybind11;
using namespace OpenXLSX;

void init_XLColor(py::module& m)
{
    py::class_<XLColor>(m, "XLColor")
        .def(py::init<>())
        .def(py::init<uint8_t, uint8_t, uint8_t, uint8_t>())
        .def(py::init<uint8_t, uint8_t, uint8_t>())
        .def(py::init<const std::string&>())
        .def("set", py::overload_cast<uint8_t, uint8_t, uint8_t, uint8_t>(&XLColor::set))
        .def("set", py::overload_cast<uint8_t, uint8_t, uint8_t>(&XLColor::set))
        .def("set", py::overload_cast<const std::string&>(&XLColor::set))
        .def("red", &XLColor::alpha)
        .def("red", &XLColor::red)
        .def("green", &XLColor::green)
        .def("blue", &XLColor::blue)
        .def("blue", &XLColor::hex);
}
