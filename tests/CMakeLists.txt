#=======================================================================================================================
# Download and define Catch2 library target
#=======================================================================================================================
add_library(Catch INTERFACE IMPORTED)
target_include_directories(Catch SYSTEM INTERFACE $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/catch/>)

#=======================================================================================================================
# Define TEST target
#=======================================================================================================================
add_executable(OpenXLSXTests EXCLUDE_FROM_ALL "")
target_sources(OpenXLSXTests
               PRIVATE
               main.cpp
               testXLDocument.cpp
               testXLCell.cpp
               testXLCellValue.cpp
               testXLCellValueProxy.cpp
               testXLFormula.cpp
               )

target_link_libraries(OpenXLSXTests
                      PRIVATE
                      OpenXLSX-static
                      Catch
                      )
target_compile_definitions(OpenXLSXTests PRIVATE OPENXLSX_STATIC_DEFINE)