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

#include <pybind11/pybind11.h>
#include <XLDocument.hpp>

namespace py = pybind11;
using namespace OpenXLSX;

void init_XLProperty(py::module &m) {
    py::enum_<XLProperty>(m, "XLProperty")
        .value("Title", XLProperty::Title)
        .value("Subject", XLProperty::Subject)
        .value("Creator", XLProperty::Creator)
        .value("Keywords", XLProperty::Keywords)
        .value("Description", XLProperty::Description)
        .value("LastModifiedBy", XLProperty::LastModifiedBy)
        .value("LastPrinted", XLProperty::LastPrinted)
        .value("CreationDate", XLProperty::CreationDate)
        .value("ModificationDate", XLProperty::ModificationDate)
        .value("Category", XLProperty::Category)
        .value("Application", XLProperty::Application)
        .value("DocSecurity", XLProperty::DocSecurity)
        .value("ScaleCrop", XLProperty::ScaleCrop)
        .value("Manager", XLProperty::Manager)
        .value("Company", XLProperty::Company)
        .value("LinksUpToDate", XLProperty::LinksUpToDate)
        .value("SharedDoc", XLProperty::SharedDoc)
        .value("HyperlinkBase", XLProperty::HyperlinkBase)
        .value("HyperlinksChanged", XLProperty::HyperlinksChanged)
        .value("AppVersion", XLProperty::AppVersion);
}

void init_XLDocument(py::module &m) {
    py::class_<XLDocument>(m, "XLDocument")
        .def(py::init<>(), "Create an empty XLDocument object.")
        .def(py::init<const std::string&>(),
             "Create an XLDocument object, and open the Excel file with the given name.",
             py::arg("filename"))
        .def("open", &XLDocument::open, "Open the Excel file with the given name.", py::arg("filename"))
        .def("create", &XLDocument::create, "Create a new Excel file.", py::arg("filename"))
        .def("close", &XLDocument::close, "Close the current Excel file.")
        .def("save", &XLDocument::save, "Save the current Excel file.")
        .def("saveAs", &XLDocument::saveAs, "Save the current Excel file with a new name.", py::arg("filename"))
        .def("name", &XLDocument::name, "Get the Excel file name.")
        .def("path", &XLDocument::path, "Get the Excel file path.")
        .def("workbook", &XLDocument::workbook, "Get the workbook object.")
        .def("property", &XLDocument::property, "Get the specified document property.", py::arg("property"))
        .def("setProperty", &XLDocument::setProperty, "Set the given property.", py::arg("property"), py::arg("getValue"))
        .def("deleteProperty", &XLDocument::deleteProperty, "Delete the given property", py::arg("property"));
}

