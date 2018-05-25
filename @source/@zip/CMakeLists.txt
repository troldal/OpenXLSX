cmake_minimum_required(VERSION 3.5)
project(OpenXLSX.Zip VERSION 0.5)

add_subdirectory(@zlib)
add_subdirectory(@libzip)

add_library(zip++ STATIC libzip++.h libzip++.cpp)
target_link_libraries(zip++ PUBLIC zip)