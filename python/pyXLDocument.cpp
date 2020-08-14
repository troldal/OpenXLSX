//
// Created by Kenneth Balslev on 14/08/2020.
//

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
        .def(py::init<>())
        .def(py::init<const std::string &>())
        .def("open", &XLDocument::open)
        .def("create", &XLDocument::create)
        .def("close", &XLDocument::close)
        .def("save", &XLDocument::save)
        .def("saveAs", &XLDocument::saveAs)
        .def("name", &XLDocument::name)
        .def("path", &XLDocument::path)
        .def("workbook", &XLDocument::workbook)
        .def("property", &XLDocument::property)
        .def("setProperty", &XLDocument::setProperty)
        .def("deleteProperty", &XLDocument::deleteProperty);
}

