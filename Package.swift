// swift-tools-version: 6.3
import PackageDescription

let package = Package(
    name: "OpenXLSX",
    platforms: [
        .macOS(.v10_15),
    ],
    products: [
        .library(name: "CxxOpenXLSX", targets: ["CxxOpenXLSX"]),
    ],
    dependencies: [
        .package(url: "https://github.com/omochi/miniz.git", branch: "swiftpm"),
        .package(url: "https://github.com/omochi/pugixml.git", branch: "swiftpm"),
    ],
    targets: [
        .target(
            name: "CxxOpenXLSX",
            dependencies: [
                .product(name: "CMiniz", package: "miniz"),
                .product(name: "CxxPugiXML", package: "pugixml"),
            ],
            path: "OpenXLSX",
            sources: [
                "sources/XLCell.cpp",
                "sources/XLCellIterator.cpp",
                "sources/XLCellRange.cpp",
                "sources/XLCellReference.cpp",
                "sources/XLCellValue.cpp",
                "sources/XLColor.cpp",
                "sources/XLColumn.cpp",
                "sources/XLComments.cpp",
                "sources/XLContentTypes.cpp",
                "sources/XLDateTime.cpp",
                "sources/XLDocument.cpp",
                "sources/XLDrawing.cpp",
                "sources/XLFormula.cpp",
                "sources/XLMergeCells.cpp",
                "sources/XLProperties.cpp",
                "sources/XLRelationships.cpp",
                "sources/XLRow.cpp",
                "sources/XLRowData.cpp",
                "sources/XLSharedStrings.cpp",
                "sources/XLSheet.cpp",
                "sources/XLStyles.cpp",
                "sources/XLTables.cpp",
                "sources/XLWorkbook.cpp",
                "sources/XLXmlData.cpp",
                "sources/XLXmlFile.cpp",
                "sources/XLXmlParser.cpp",
                "sources/XLZipArchive.cpp",
            ],
            publicHeadersPath: ".",
            cxxSettings: [
                .headerSearchPath("headers"),
            ]
        ),
        .target(
            name: "CxxOpenXLSXTestSupport",
            dependencies: ["CxxOpenXLSX"],
            path: "SwiftPMTests/CxxOpenXLSXTestSupport",
            publicHeadersPath: "include"
        ),
        .testTarget(
            name: "CxxOpenXLSXTests",
            dependencies: [
                "CxxOpenXLSX",
                "CxxOpenXLSXTestSupport",
            ],
            path: "SwiftPMTests/CxxOpenXLSXTests",
            resources: [
                .process("Resources"),
            ],
            swiftSettings: [
                .interoperabilityMode(.Cxx),
            ]
        ),
    ],
    cxxLanguageStandard: .cxx17
)
