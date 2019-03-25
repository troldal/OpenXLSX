#=======================================================================================================================
# Define Benchmark target
#=======================================================================================================================
add_executable(Benchmark Benchmark.cpp table_printer.h)
target_link_libraries(Benchmark PUBLIC OpenXLSX)
target_include_directories(Benchmark PUBLIC ${OPENXLSX_INSTALLDIR}/include)
set_target_properties(Benchmark PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${OPENXLSX_INSTALLDIR}/bin)
