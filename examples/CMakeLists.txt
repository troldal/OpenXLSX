add_subdirectory(external/nowide EXCLUDE_FROM_ALL)

#=======================================================================================================================
# Define Demo1 target
#=======================================================================================================================
add_executable(Demo1 Demo1.cpp)
target_link_libraries(Demo1 PRIVATE OpenXLSX-static)
target_compile_definitions(Demo1 PRIVATE OPENXLSX_STATIC_DEFINE)

#=======================================================================================================================
# Define Demo2 target
#=======================================================================================================================
add_executable(Demo2 Demo2.cpp)
target_link_libraries(Demo2 PRIVATE OpenXLSX-static)
target_compile_definitions(Demo2 PRIVATE OPENXLSX_STATIC_DEFINE)

#=======================================================================================================================
# Define Demo3 target
#=======================================================================================================================
add_executable(Demo3 Demo3.cpp)
target_link_libraries(Demo3 PRIVATE OpenXLSX-static nowide::nowide)
target_compile_definitions(Demo3 PRIVATE OPENXLSX_STATIC_DEFINE)

#=======================================================================================================================
# Define Demo4 target
#=======================================================================================================================
add_executable(Demo4 Demo4.cpp)
target_link_libraries(Demo4 PRIVATE OpenXLSX-static)
target_compile_definitions(Demo4 PRIVATE OPENXLSX_STATIC_DEFINE)

#=======================================================================================================================
# Define Demo5 target
#=======================================================================================================================
add_executable(Demo5 Demo5.cpp)
target_link_libraries(Demo5 PRIVATE OpenXLSX-static)
target_compile_definitions(Demo5 PRIVATE OPENXLSX_STATIC_DEFINE)

#=======================================================================================================================
# Define Demo6 target
#=======================================================================================================================
add_executable(Demo6 Demo6.cpp)
target_link_libraries(Demo6 PRIVATE OpenXLSX-static)
target_compile_definitions(Demo6 PRIVATE OPENXLSX_STATIC_DEFINE)

#=======================================================================================================================
# Define Demo7 target
#=======================================================================================================================
add_executable(Demo7 Demo7.cpp)
target_link_libraries(Demo7 PRIVATE OpenXLSX-static)
target_compile_definitions(Demo7 PRIVATE OPENXLSX_STATIC_DEFINE)

#=======================================================================================================================
# Define Demo8 target
#=======================================================================================================================
add_executable(Demo8 Demo8.cpp)
target_link_libraries(Demo8 PRIVATE OpenXLSX-static)
target_compile_definitions(Demo8 PRIVATE OPENXLSX_STATIC_DEFINE)

#=======================================================================================================================
# Define Demo9 target
#=======================================================================================================================
add_executable(Demo9 Demo9.cpp)
target_link_libraries(Demo9 PRIVATE OpenXLSX-static)
target_compile_definitions(Demo9 PRIVATE OPENXLSX_STATIC_DEFINE)