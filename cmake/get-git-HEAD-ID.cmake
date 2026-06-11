# Tool Origin: https://github.com/shzycat/Use-CMake-to-get-Git-HEAD-info
# Author: https://github.com/shzycat/
# Purpose:
#   Use CMake to get current git repo's HEAD commit id and time.
#   Then as preprocessor macros, they can be used in your code.
#   Every time you use CMake as your build tools,
#    the HEAD info is "Hard-Coding" into your executable or library,
#    may sometimes be very helpful,
#    eg. you want to directly know this .so file's code version.

# reach the root dir of current git repository
execute_process(
  COMMAND git rev-parse --show-toplevel
  WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
  OUTPUT_VARIABLE GIT_REPO_ROOT
  OUTPUT_STRIP_TRAILING_WHITESPACE
)

# 2026-03-22 LU: added check that git / GIT_REPO_ROOT are actually available
if(EXISTS "${GIT_REPO_ROOT}")
    # get the HEAD's commit ID
    # if you want commit id longer than 16, change the number following.
    execute_process(
    # COMMAND git rev-parse --short=16 HEAD
    COMMAND git rev-parse --short=8 HEAD
    WORKING_DIRECTORY ${GIT_REPO_ROOT}
    OUTPUT_VARIABLE GIT_COMMIT_ID
    OUTPUT_STRIP_TRAILING_WHITESPACE
    )
    # parse the HEAD's time as string
    execute_process(
    COMMAND git show -s --format=%ci ${GIT_COMMIT_ID}
    WORKING_DIRECTORY ${GIT_REPO_ROOT}
    OUTPUT_VARIABLE GIT_COMMIT_ID_TIME
    OUTPUT_STRIP_TRAILING_WHITESPACE
    )
else()
    set(GIT_COMMIT_ID "(not available)")
    set(GIT_COMMIT_ID_TIME "(not available)")
endif()

# define the commit ID and time as macros,
# so they can be used in your executable.(aka. target in CMake)
add_definitions(-DGIT_COMMIT_ID="${GIT_COMMIT_ID}")
add_definitions(-DGIT_COMMIT_ID_TIME="${GIT_COMMIT_ID_TIME}")
