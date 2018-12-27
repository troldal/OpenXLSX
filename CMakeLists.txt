#=======================================================================================================================
# Preamble
#=======================================================================================================================
cmake_minimum_required(VERSION 3.11 FATAL_ERROR)
project (OpenXLSX.All)

#=======================================================================================================================
# Set C/C++ compiler version
#=======================================================================================================================
set(CMAKE_C_STANDARD 99)
set(CMAKE_C_STANDARD_REQUIRED ON)
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

add_subdirectory(@library)
add_subdirectory(@benchmark)
add_subdirectory(@demo)