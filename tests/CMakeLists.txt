#=======================================================================================================================
# Download and define Catch2 library target
#=======================================================================================================================
add_library(Catch INTERFACE IMPORTED)
target_include_directories(Catch SYSTEM INTERFACE $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/catch/>)

#=======================================================================================================================
# Define TEST target
#=======================================================================================================================
add_executable(OpenXLSXTests EXCLUDE_FROM_ALL "")
target_sources(OpenXLSXTests
               PRIVATE
               mainTest.cpp
               documentTest.cpp
               )

target_link_libraries(OpenXLSXTests
                      PRIVATE
                      OpenXLSX-shared
                      Catch
                      )

#=======================================================================================================================
# Enable code coverage
#=======================================================================================================================
#if (ENABLE_COVERAGE AND CMAKE_BUILD_TYPE STREQUAL "Debug")
#    if ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Clang" AND NOT WIN32)
#        target_compile_options(OpenXLSXTests PRIVATE -fprofile-instr-generate -fcoverage-mapping)
#        target_link_options(OpenXLSXTests PRIVATE -fprofile-instr-generate)
#        #Uncomment in case of linker errors
#        #link_libraries(clang_rt.profile-x86_64)
#    elseif ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU")
#        target_compile_options(OpenXLSXTests PRIVATE --coverage)
#        target_link_options(OpenXLSXTests PRIVATE --coverage)
#        #Uncomment in case of linker errors
#        #link_libraries(gcov)
#    endif ()
#endif ()