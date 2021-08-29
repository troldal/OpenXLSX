# Project testing to consume Nowide.Nowide
cmake_minimum_required(VERSION 3.8.0)
project(nowide_example_project)

find_package(nowide 11.0 REQUIRED)

add_executable(example example_main.cpp)
target_link_libraries(example PRIVATE nowide::nowide)
enable_testing()
add_test(NAME example COMMAND example)
