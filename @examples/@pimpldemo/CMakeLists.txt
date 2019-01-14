#=======================================================================================================================
# Define Demo target
#=======================================================================================================================

add_executable(PimplDemo Demo.cpp)
target_link_libraries(PimplDemo PUBLIC -lOpenXLSX)
target_link_directories(PimplDemo PUBLIC ${OPENXLSX_INSTALLDIR}/lib)
target_link_directories(PimplDemo PUBLIC ${OPENXLSX_INSTALLDIR}/bin)
target_include_directories(PimplDemo PUBLIC ${CMAKE_CURRENT_LIST_DIR}/../../@library/@openxlsx/interface/headers)

set_target_properties(PimplDemo PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${OPENXLSX_INSTALLDIR}/bin)
