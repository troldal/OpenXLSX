//
// Created by Kenneth Balslev on 14/08/2020.
//

#include <pybind11/pybind11.h>

namespace py = pybind11;

void init_XLCell(py::module& m);
void init_XLCellRange(py::module& m);
void init_XLCellReference(py::module& m);
void init_XLCellValue(py::module& m);
void init_XLChartsheet(py::module& m);
void init_XLColor(py::module& m);
void init_XLColumn(py::module& m);
void init_XLDocument(py::module& m);
void init_XLProperty(py::module& m);
void init_XLRow(py::module& m);
void init_XLSheet(py::module& m);
void init_XLSheetState(py::module& m);
void init_XLSheetType(py::module& m);
void init_XLValueType(py::module& m);
void init_XLWorkbook(py::module& m);
void init_XLWorksheet(py::module& m);

PYBIND11_MODULE(OpenXLSX, m)
{
    init_XLCell(m);
    init_XLCellRange(m);
    init_XLCellReference(m);
    init_XLCellValue(m);
    init_XLChartsheet(m);
    init_XLColor(m);
    init_XLColumn(m);
    init_XLDocument(m);
    init_XLProperty(m);
    init_XLRow(m);
    init_XLSheet(m);
    init_XLSheetState(m);
    init_XLSheetType(m);
    init_XLValueType(m);
    init_XLWorkbook(m);
    init_XLWorksheet(m);
}