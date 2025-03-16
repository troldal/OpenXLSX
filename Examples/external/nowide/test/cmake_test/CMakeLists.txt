# Copyright 2021 Alexander Grund
# Distributed under the Nowide Software License, Version 1.0.
# See accompanying file LICENSE_1_0.txt or copy at http://www.boost.org/LICENSE_1_0.txt

cmake_minimum_required(VERSION 3.15 FATAL_ERROR)

project(cmake_subdir_test LANGUAGES CXX)

    find_package(nowide 11 REQUIRED)

add_executable(main main.cpp)
target_link_libraries(main nowide::nowide)

enable_testing()
add_test(NAME main COMMAND main)
