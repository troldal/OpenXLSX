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

#=======================================================================================================================
# Define Demo3 target
#=======================================================================================================================
add_executable(Demo3 Demo3.cpp)
target_link_libraries(Demo3 PUBLIC OpenXLSX)
target_include_directories(Demo3 PUBLIC ${OPENXLSX_INSTALLDIR}/include)
set_target_properties(Demo3 PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${OPENXLSX_INSTALLDIR}/bin)

#=======================================================================================================================
# Define Demo4 target
#=======================================================================================================================
add_executable(Demo4 Demo4.cpp)
target_link_libraries(Demo4 PUBLIC OpenXLSX)
target_include_directories(Demo4 PUBLIC ${OPENXLSX_INSTALLDIR}/include)
set_target_properties(Demo4 PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${OPENXLSX_INSTALLDIR}/bin)

#=======================================================================================================================
# Define Demo5 target
#=======================================================================================================================
add_executable(Demo5 Demo5.cpp)
target_link_libraries(Demo5 PUBLIC OpenXLSX)
target_include_directories(Demo5 PUBLIC ${OPENXLSX_INSTALLDIR}/include)
set_target_properties(Demo5 PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${OPENXLSX_INSTALLDIR}/bin)

#=======================================================================================================================
# Define Demo5 target
#=======================================================================================================================
add_executable(Demo6 Demo6.cpp)
target_link_libraries(Demo6 PUBLIC OpenXLSX)
target_include_directories(Demo6 PUBLIC ${OPENXLSX_INSTALLDIR}/include)
set_target_properties(Demo6 PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${OPENXLSX_INSTALLDIR}/bin)