include(FetchContent)
FetchContent_Declare(
        catch
        GIT_REPOSITORY https://github.com/catchorg/Catch2.git
        GIT_TAG c1720d0
)
FetchContent_GetProperties(catch)

if(NOT catch_POPULATED)
    FetchContent_Populate(catch)
    add_subdirectory(
            ${catch_SOURCE_DIR}
            ${catch_BINARY_DIR}
    )
endif()

add_executable(InterfaceTest "")
target_sources(InterfaceTest PRIVATE
        Test00_Main.cpp
        Test01_DocumentCreation.cpp
        Test02_DocumentProperties.cpp
        Test03_Workbook.cpp
        Test04-06_Sheet.cpp
        Test07_CellReference.cpp
        Test08_Cell.cpp)

target_link_libraries(InterfaceTest INTERFACE Catch2)
target_link_libraries(InterfaceTest PUBLIC OpenXLSX)
#target_link_directories(InterfaceTest PUBLIC ${OPENXLSX_INSTALLDIR}/lib)
#target_link_directories(InterfaceTest PUBLIC ${OPENXLSX_INSTALLDIR}/bin)
target_include_directories(InterfaceTest PRIVATE ${catch_SOURCE_DIR}/single_include/catch2)
#target_include_directories(InterfaceTest PUBLIC ${CMAKE_CURRENT_LIST_DIR}/../../@library/@openxlsx/interface/headers)