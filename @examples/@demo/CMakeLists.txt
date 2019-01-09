#=======================================================================================================================
# Define Demo target
#=======================================================================================================================

add_executable(Demo main.cpp)
target_link_libraries(Demo PUBLIC -lOpenXLSX)
target_link_directories(Demo PUBLIC ${OPENXLSX_INSTALLDIR}/lib)
target_link_directories(Demo PUBLIC ${OPENXLSX_INSTALLDIR}/bin)
target_include_directories(Demo PUBLIC ${OPENXLSX_INSTALLDIR}/include)

set_target_properties(Demo PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${OPENXLSX_INSTALLDIR}/bin)
