#=======================================================================================================================
# Download and define Catch2 library target
#=======================================================================================================================
file(MAKE_DIRECTORY ${CMAKE_CURRENT_BINARY_DIR}/_deps/Catch2)
if (NOT EXISTS ${CMAKE_CURRENT_BINARY_DIR}/_deps/Catch2/catch.hpp)
    message(STATUS "Catch2 header file not found. Downloading...")
    file(DOWNLOAD
         https://raw.githubusercontent.com/catchorg/Catch2/master/single_include/catch2/catch.hpp
         ${CMAKE_CURRENT_BINARY_DIR}/_deps/Catch2/catch.hpp
         TIMEOUT 60)
endif ()

add_library(Catch INTERFACE)
target_include_directories(Catch SYSTEM INTERFACE ${CMAKE_CURRENT_BINARY_DIR}/_deps/Catch2)

#=======================================================================================================================
# Define TEST target
#=======================================================================================================================
add_executable(InterfaceTest EXCLUDE_FROM_ALL "")
target_sources(InterfaceTest
               PRIVATE
               Test00_Main.cpp
               Test01_DocumentCreation.cpp
               Test02_DocumentProperties.cpp
               Test03_Workbook.cpp
               Test04-06_Sheet.cpp
               Test07_CellReference.cpp
               Test08_Cell.cpp
               )

target_link_libraries(InterfaceTest
                      PRIVATE
                      OpenXLSX-objs
                      Catch
                      )

#=======================================================================================================================
# Enable code coverage
#=======================================================================================================================
if (ENABLE_COVERAGE AND CMAKE_BUILD_TYPE STREQUAL "Debug")
    if ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Clang" AND NOT WIN32)
        target_compile_options(InterfaceTest PRIVATE -fprofile-instr-generate -fcoverage-mapping)
        target_link_options(InterfaceTest PRIVATE -fprofile-instr-generate)
        #Uncomment in case of linker errors
        #link_libraries(clang_rt.profile-x86_64)
    elseif ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU")
        target_compile_options(InterfaceTest PRIVATE --coverage)
        target_link_options(InterfaceTest PRIVATE --coverage)
        #Uncomment in case of linker errors
        #link_libraries(gcov)
    endif ()
endif ()