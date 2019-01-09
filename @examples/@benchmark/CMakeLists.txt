#=======================================================================================================================
# Define Benchmark target
#=======================================================================================================================

add_executable(Benchmark benchmark.cpp table_printer.h)
target_link_libraries(Benchmark PUBLIC -lOpenXLSX)
target_link_directories(Benchmark PUBLIC ${OPENXLSX_INSTALLDIR}/lib)
target_link_directories(Benchmark PUBLIC ${OPENXLSX_INSTALLDIR}/bin)
target_include_directories(Benchmark PUBLIC ${OPENXLSX_INSTALLDIR}/include)

set_target_properties(Benchmark PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${OPENXLSX_INSTALLDIR}/bin)
