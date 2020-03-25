#=======================================================================================================================
# Define Benchmark target
#=======================================================================================================================
add_executable(Benchmark Benchmark.cpp table_printer.h)
target_link_libraries(Benchmark PUBLIC OpenXLSX-static)
set_target_properties(Benchmark PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${OPENXLSX_INSTALLDIR}/bin)
#add_custom_command(TARGET Benchmark POST_BUILD COMMAND Benchmark)
