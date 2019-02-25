#=======================================================================================================================
# Define Demo1 target
#=======================================================================================================================

add_executable(Demo1 Demo1.cpp)
target_link_libraries(Demo1 PUBLIC -lOpenXLSX)
target_link_directories(Demo1 PUBLIC ${OPENXLSX_INSTALLDIR}/lib)
target_link_directories(Demo1 PUBLIC ${OPENXLSX_INSTALLDIR}/bin)
target_include_directories(Demo1 PUBLIC ${OPENXLSX_INSTALLDIR}/include)

set_target_properties(Demo1 PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${OPENXLSX_INSTALLDIR}/bin)


#=======================================================================================================================
# Define Demo2 target
#=======================================================================================================================

add_executable(Demo2 Demo2.cpp)
target_link_libraries(Demo2 PUBLIC -lOpenXLSX)
target_link_directories(Demo2 PUBLIC ${OPENXLSX_INSTALLDIR}/lib)
target_link_directories(Demo2 PUBLIC ${OPENXLSX_INSTALLDIR}/bin)
target_include_directories(Demo2 PUBLIC ${OPENXLSX_INSTALLDIR}/include)

set_target_properties(Demo2 PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${OPENXLSX_INSTALLDIR}/bin)
add_dependencies(Demo2 Demo1)

