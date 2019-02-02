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

add_executable(CInterfaceTest "")
target_sources(CInterfaceTest PRIVATE
        Test00_Main.cpp
        Test01_DocumentCreation.cpp
        Test02_DocumentProperties.cpp)

target_link_libraries(CInterfaceTest INTERFACE Catch2)
target_link_libraries(CInterfaceTest PUBLIC OpenXLSX)
#target_link_directories(InterfaceTest PUBLIC ${OPENXLSX_INSTALLDIR}/lib)
#target_link_directories(InterfaceTest PUBLIC ${OPENXLSX_INSTALLDIR}/bin)
target_include_directories(CInterfaceTest PRIVATE ${catch_SOURCE_DIR}/single_include/catch2)
#target_include_directories(InterfaceTest PUBLIC ${CMAKE_CURRENT_LIST_DIR}/../../@library/@openxlsx/interfaces/headers)