include(FetchContent)
FetchContent_Declare(
        catch
        GIT_REPOSITORY https://github.com/catchorg/Catch2.git
        GIT_TAG c1720d0
)
FetchContent_GetProperties(catch)

if(NOT catch_POPULATED)
    FetchContent_Populate(catch)
    add_subdirectory(
            ${catch_SOURCE_DIR}
            ${catch_BINARY_DIR}
    )
endif()

add_executable(Test2 "")
target_sources(Test2 PRIVATE Test2.cpp)
target_link_libraries(Test2 INTERFACE Catch2)
target_include_directories(Test2 PRIVATE ${catch_SOURCE_DIR}/single_include/catch2)

