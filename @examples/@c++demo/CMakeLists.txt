#=======================================================================================================================
# Define Demo target
#=======================================================================================================================

add_executable(CxxDemo Demo.cpp)
target_link_libraries(CxxDemo PUBLIC -lOpenXLSX)
target_link_directories(CxxDemo PUBLIC ${OPENXLSX_INSTALLDIR}/lib)
target_link_directories(CxxDemo PUBLIC ${OPENXLSX_INSTALLDIR}/bin)
target_include_directories(CxxDemo PUBLIC ${OPENXLSX_INSTALLDIR}/include)

set_target_properties(CxxDemo PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${OPENXLSX_INSTALLDIR}/bin)
