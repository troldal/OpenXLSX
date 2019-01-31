#=======================================================================================================================
# Define Demo target
#=======================================================================================================================

add_executable(CDemo Demo.c)
target_link_libraries(CDemo PUBLIC -lOpenXLSX)
target_link_directories(CDemo PUBLIC ${OPENXLSX_INSTALLDIR}/lib)
target_link_directories(CDemo PUBLIC ${OPENXLSX_INSTALLDIR}/bin)
target_include_directories(CDemo PUBLIC ${OPENXLSX_INSTALLDIR}/include)

set_target_properties(CDemo PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${OPENXLSX_INSTALLDIR}/bin)
