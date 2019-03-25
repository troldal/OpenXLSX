add_executable(InterfaceTest "")
target_sources(InterfaceTest PRIVATE
        Test00_Main.cpp
        Test01_DocumentCreation.cpp
        Test02_DocumentProperties.cpp
        Test03_Workbook.cpp
        Test04-06_Sheet.cpp
        Test07_CellReference.cpp
        Test08_Cell.cpp)

target_link_libraries(InterfaceTest PRIVATE Catch INTERFACE Catch)
target_link_libraries(InterfaceTest PUBLIC OpenXLSX)