include(FetchContent)
FetchContent_Declare(
        googletest
        GIT_REPOSITORY https://github.com/google/googletest.git
        GIT_TAG release-1.8.0
)
FetchContent_GetProperties(googletest)

if(NOT googletest_POPULATED)
    FetchContent_Populate(googletest)
    # Prevent GoogleTest from overriding our compiler/linker options
    # when building with Visual Studio
    set(gtest_force_shared_crt ON CACHE BOOL "" FORCE)
    # Prevent GoogleTest from using PThreads
    set(gtest_disable_pthreads ON CACHE BOOL "" FORCE)
    # adds the targers: gtest, gtest_main, gmock, gmock_main
    add_subdirectory(
            ${googletest_SOURCE_DIR}
            ${googletest_BINARY_DIR}
    )
    # Silence std::tr1 warning on MSVC
    if(MSVC)
        foreach(_tgt gtest gtest_main gmock gmock_main)
            target_compile_definitions(${_tgt}
                    PRIVATE
                    "_SILENCE_TR1_NAMESPACE_DEPRECATION_WARNING"
                    )
        endforeach()
    endif()
endif()

add_executable(Test "")
target_sources(Test PRIVATE Test.cpp)
target_link_libraries(Test PRIVATE gtest_main)
target_include_directories(Test PRIVATE ${googletest_SOURCE_DIR}/googletest/include)
