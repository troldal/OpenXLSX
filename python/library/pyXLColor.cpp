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

  Copyright (c) 2020, Kenneth Troldal Balslev

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

#include <XLColor.hpp>
#include <pybind11/pybind11.h>

namespace py = pybind11;
using namespace OpenXLSX;

void init_XLColor(py::module& m)
{
    py::class_<XLColor>(m, "XLColor")
        .def(py::init<>(), "Create an empty color object.")
        .def(py::init<uint8_t, uint8_t, uint8_t, uint8_t>(),
             "Create a color object with the given values.",
             py::arg("alpha"),
             py::arg("red"),
             py::arg("green"),
             py::arg("blue"))
        .def(py::init<uint8_t, uint8_t, uint8_t>(),
             "Create a color object with the given values.",
             py::arg("red"),
             py::arg("green"),
             py::arg("blue"))
        .def(py::init<const std::string&>(), "Create a color object with the given hex code.", py::arg("hexCode"))
        .def("set",
             py::overload_cast<uint8_t, uint8_t, uint8_t, uint8_t>(&XLColor::set),
             "Set the color with the given values.",
             py::arg("alpha"),
             py::arg("red"),
             py::arg("green"),
             py::arg("blue"))
        .def("set",
             py::overload_cast<uint8_t, uint8_t, uint8_t>(&XLColor::set),
             "Set the color with the given values.",
             py::arg("red"),
             py::arg("green"),
             py::arg("blue"))
        .def("set", py::overload_cast<const std::string&>(&XLColor::set), "Set the color with the given hex code.", py::arg("hexCode"))
        .def("red", &XLColor::alpha, "Get the alpha getValue of the color.")
        .def("red", &XLColor::red, "Get the red color getValue for the color object.")
        .def("green", &XLColor::green, "Get the green color getValue for the color object.")
        .def("blue", &XLColor::blue, "Get the blue color getValue for the color object.")
        .def("blue", &XLColor::hex, "Get the hex code for the color object.");
}
