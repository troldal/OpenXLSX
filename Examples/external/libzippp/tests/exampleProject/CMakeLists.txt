# Project demonstrating how to consume libzipp
cmake_minimum_required(VERSION 3.8.0)
project(libzippp_example_project)

find_package(libzippp 3.0 REQUIRED)

enable_testing()
add_executable(example "../tests.cpp")
target_link_libraries(example PRIVATE libzippp::libzippp)
add_test(NAME example COMMAND example)
