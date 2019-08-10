#=======================================================================================================================
# Define Demo1 target
#=======================================================================================================================
add_executable(Demo1 Demo1.cpp)
target_link_libraries(Demo1 PUBLIC OpenXLSX)
target_include_directories(Demo1 PUBLIC ${OPENXLSX_INSTALLDIR}/include)
set_target_properties(Demo1 PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${OPENXLSX_INSTALLDIR}/bin)

#=======================================================================================================================
# Define Demo2 target
#=======================================================================================================================
add_executable(Demo2 Demo2.cpp)
target_link_libraries(Demo2 PUBLIC OpenXLSX)
target_include_directories(Demo2 PUBLIC ${OPENXLSX_INSTALLDIR}/include)
set_target_properties(Demo2 PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${OPENXLSX_INSTALLDIR}/bin)


