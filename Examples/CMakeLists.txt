#=======================================================================================================================
# Preamble
#=======================================================================================================================
cmake_minimum_required(VERSION 3.15 FATAL_ERROR)
project(OpenXLSX.Examples)
list(APPEND CMAKE_MODULE_PATH "${CMAKE_CURRENT_SOURCE_DIR}/cmake")

find_package(LIBZIP MODULE REQUIRED)

#=======================================================================================================================
# Include boost::nowide library (required for Demo3)
#=======================================================================================================================
add_subdirectory(external/nowide EXCLUDE_FROM_ALL)
#add_subdirectory(external/libzippp EXCLUDE_FROM_ALL)

#=======================================================================================================================
# Define Demo1 target
#=======================================================================================================================
add_executable(Demo1 Demo1.cpp)
target_link_libraries(Demo1 PRIVATE OpenXLSX::OpenXLSX nowide::nowide libzip::zip)

#=======================================================================================================================
# Define Demo2 target
#=======================================================================================================================
add_executable(Demo2 Demo2.cpp)
target_link_libraries(Demo2 PRIVATE OpenXLSX::OpenXLSX nowide::nowide)

#=======================================================================================================================
# Define Demo3 target
#=======================================================================================================================
add_executable(Demo3 Demo3.cpp)
target_link_libraries(Demo3 PRIVATE OpenXLSX::OpenXLSX nowide::nowide libzip::zip)

#=======================================================================================================================
# Define Demo4 target
#=======================================================================================================================
add_executable(Demo4 Demo4.cpp)
target_link_libraries(Demo4 PRIVATE OpenXLSX::OpenXLSX nowide::nowide libzip::zip)

#=======================================================================================================================
# Define Demo9 target
#=======================================================================================================================
add_executable(Demo5 Demo5.cpp)
target_link_libraries(Demo5 PRIVATE OpenXLSX::OpenXLSX nowide::nowide libzip::zip)

#=======================================================================================================================
# Define Demo6 target
#=======================================================================================================================
add_executable(Demo6 Demo6.cpp)
target_link_libraries(Demo6 PRIVATE OpenXLSX::OpenXLSX nowide::nowide libzip::zip)

#=======================================================================================================================
# Define Demo7 target
#=======================================================================================================================
add_executable(Demo7 Demo7.cpp)
target_link_libraries(Demo7 PRIVATE OpenXLSX::OpenXLSX nowide::nowide libzip::zip)

#=======================================================================================================================
# Define Demo8 target
#=======================================================================================================================
add_executable(Demo8 Demo8.cpp)
target_link_libraries(Demo8 PRIVATE OpenXLSX::OpenXLSX nowide::nowide libzip::zip)


