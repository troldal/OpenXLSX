import Foundation
import CxxOpenXLSXTestSupport
import Testing

@Test func readsCellValuesFromResourceWorkbook() throws {
    let workbookURL = try #require(Bundle.module.url(forResource: "simple", withExtension: "xlsx"))

    #expect(readCellString(workbookPath: workbookURL.path, cellReference: "A1") == "A")
    #expect(readCellString(workbookPath: workbookURL.path, cellReference: "B1") == "B")
    #expect(readCellString(workbookPath: workbookURL.path, cellReference: "C1") == "C")
    #expect(readCellString(workbookPath: workbookURL.path, cellReference: "A2") == "1")
    #expect(readCellString(workbookPath: workbookURL.path, cellReference: "B2") == "2")
    #expect(readCellString(workbookPath: workbookURL.path, cellReference: "C2") == "3")
}

@Test func readsExtListFromResourceWorkbook() throws {
    let workbookURL = try #require(Bundle.module.url(forResource: "simple", withExtension: "xlsx"))

    #expect(readExtList(workbookPath: workbookURL.path) == "<extLst />\n")
}

private func readCellString(workbookPath: String, cellReference: String) -> String {
    String(cxx_readCellString(
        std.string(workbookPath),
        std.string(cellReference)
    ))
}

private func readExtList(workbookPath: String) -> String {
    String(cxx_readExtList(std.string(workbookPath)))
}
