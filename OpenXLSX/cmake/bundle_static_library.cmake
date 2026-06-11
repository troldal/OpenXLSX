# MIT License
# 
# Copyright (c) 2019 Cristian Adam
# 
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

function(bundle_static_library tgt_name bundled_tgt_name)
  list(APPEND static_libs ${tgt_name})

  function(_recursively_collect_dependencies input_target)
    set(_input_link_libraries LINK_LIBRARIES)
    get_target_property(_input_type ${input_target} TYPE)
    if (${_input_type} STREQUAL "INTERFACE_LIBRARY")
      set(_input_link_libraries INTERFACE_LINK_LIBRARIES)
    endif()
    get_target_property(public_dependencies ${input_target} ${_input_link_libraries})
    foreach(dependency IN LISTS public_dependencies)
      if(TARGET ${dependency})
        get_target_property(alias ${dependency} ALIASED_TARGET)
        if (TARGET ${alias})
          set(dependency ${alias})
        endif()
        get_target_property(_type ${dependency} TYPE)
        if (${_type} STREQUAL "STATIC_LIBRARY")
          list(APPEND static_libs ${dependency})
        endif()

        get_property(library_already_added
          GLOBAL PROPERTY _${tgt_name}_static_bundle_${dependency})
        if (NOT library_already_added)
          set_property(GLOBAL PROPERTY _${tgt_name}_static_bundle_${dependency} ON)
          _recursively_collect_dependencies(${dependency})
        endif()
      endif()
    endforeach()
    set(static_libs ${static_libs} PARENT_SCOPE)
  endfunction()

  _recursively_collect_dependencies(${tgt_name})

  list(REMOVE_DUPLICATES static_libs)

  set(bundled_tgt_full_name 
    ${CMAKE_BINARY_DIR}/${CMAKE_STATIC_LIBRARY_PREFIX}${bundled_tgt_name}${CMAKE_STATIC_LIBRARY_SUFFIX})

  if (CMAKE_CXX_COMPILER_ID MATCHES "^(Clang|GNU)$")
    file(WRITE ${CMAKE_BINARY_DIR}/${bundled_tgt_name}.ar.in
      "CREATE ${bundled_tgt_full_name}\n" )
        
    foreach(tgt IN LISTS static_libs)
      file(APPEND ${CMAKE_BINARY_DIR}/${bundled_tgt_name}.ar.in
        "ADDLIB $<TARGET_FILE:${tgt}>\n")
    endforeach()
    
    file(APPEND ${CMAKE_BINARY_DIR}/${bundled_tgt_name}.ar.in "SAVE\n")
    file(APPEND ${CMAKE_BINARY_DIR}/${bundled_tgt_name}.ar.in "END\n")

    file(GENERATE
      OUTPUT ${CMAKE_BINARY_DIR}/${bundled_tgt_name}.ar
      INPUT ${CMAKE_BINARY_DIR}/${bundled_tgt_name}.ar.in)

    set(ar_tool ${CMAKE_AR})
    if (CMAKE_INTERPROCEDURAL_OPTIMIZATION)
      set(ar_tool ${CMAKE_CXX_COMPILER_AR})
    endif()

    add_custom_command(
      COMMAND ${ar_tool} -M < ${CMAKE_BINARY_DIR}/${bundled_tgt_name}.ar
      OUTPUT ${bundled_tgt_full_name}
      COMMENT "Bundling ${bundled_tgt_name}"
      VERBATIM)
  elseif(MSVC)
    find_program(lib_tool lib)

    foreach(tgt IN LISTS static_libs)
      list(APPEND static_libs_full_names $<TARGET_FILE:${tgt}>)
    endforeach()

    add_custom_command(
      COMMAND ${lib_tool} /NOLOGO /OUT:${bundled_tgt_full_name} ${static_libs_full_names}
      OUTPUT ${bundled_tgt_full_name}
      COMMENT "Bundling ${bundled_tgt_name}"
      VERBATIM)
  else()
    message(FATAL_ERROR "Unknown bundle scenario!")
  endif()

  add_custom_target(bundling_target ALL DEPENDS ${bundled_tgt_full_name})
  add_dependencies(bundling_target ${tgt_name})

  add_library(${bundled_tgt_name} STATIC IMPORTED)
  set_target_properties(${bundled_tgt_name} 
    PROPERTIES 
      IMPORTED_LOCATION ${bundled_tgt_full_name}
      INTERFACE_INCLUDE_DIRECTORIES $<TARGET_PROPERTY:${tgt_name},INTERFACE_INCLUDE_DIRECTORIES>)
  add_dependencies(${bundled_tgt_name} bundling_target)

endfunction()


#=======================================================================================================================
# BEGIN OF FUNCTION: bundle_static_library_simplified
# Purpose: replace ${CMAKE_BINARY_DIR} with ${CMAKE_LIBRARY_OUTPUT_DIRECTORY} and process a *known* list of dependencies
#            without the need for the function to _recursively_collect_dependencies
#          Also: better understand the function through learning by doing
# Status: 2026-03-24 - static_libs_system is no longer in use, static_libs must contain full paths (generator expressions if needed) to static targets
# Parameters:
#   lib_name            original library name for which dependencies shall be statically linked
#   lib_name_combined   output library name that includes the dependencies
#   static_libs         list of dependencies to link into lib_name_combined
#   static_libs_system  list of fully resolved dependencies (of preinstalled libraries) to link into lib_name_combined
#=======================================================================================================================
function(bundle_static_library_simplified lib_name lib_name_combined static_libs static_libs_system)
    set(full_lib_name_combined "${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/${CMAKE_STATIC_LIBRARY_PREFIX}${lib_name_combined}${CMAKE_STATIC_LIBRARY_SUFFIX}")

    if (CMAKE_CXX_COMPILER_ID MATCHES "^(Clang|GNU)$")
        file(WRITE ${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/${lib_name_combined}.ar.in
          "CREATE ${full_lib_name_combined}\n" )

        foreach(tgt IN LISTS static_libs)
            message( NOTICE "static target is ${tgt}" )
            file(APPEND ${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/${lib_name_combined}.ar.in
                    "ADDLIB ${tgt}\n")
                    # "ADDLIB $<TARGET_FILE:${tgt}>\n")
        endforeach()
        foreach(tgt IN LISTS static_libs_system)
            message( NOTICE "installed static target is ${tgt}" )
            file(APPEND ${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/${lib_name_combined}.ar.in
                    "ADDLIB ${tgt}>\n")
        endforeach()

        file(APPEND ${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/${lib_name_combined}.ar.in "SAVE\n")
        file(APPEND ${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/${lib_name_combined}.ar.in "END\n")

        file(GENERATE
             OUTPUT ${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/${lib_name_combined}.ar
             INPUT ${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/${lib_name_combined}.ar.in)

        set(ar_tool ${CMAKE_AR})
        if (CMAKE_INTERPROCEDURAL_OPTIMIZATION)
            set(ar_tool ${CMAKE_CXX_COMPILER_AR})
        endif()

        add_custom_command(
            COMMAND ${ar_tool} -M < ${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/${lib_name_combined}.ar
            OUTPUT ${full_lib_name_combined}
            COMMENT "Bundling ${lib_name_combined}"
            VERBATIM)
    elseif(MSVC) # untested, unsupported ATM
        message( FATAL_ERROR "OPENXLSX_BUNDLED_STATIC_LIB is not yet supported for MSVC" )

        find_program(lib_tool lib)

        # foreach(tgt IN LISTS static_libs)
        #     list(APPEND static_libs_full_names $<TARGET_FILE:${tgt}>)
        # endforeach()

        add_custom_command(
            # COMMAND ${lib_tool} /NOLOGO /OUT:${full_lib_name_combined} ${static_libs_full_names}
            COMMAND ${lib_tool} /NOLOGO /OUT:${full_lib_name_combined} ${static_libs}
            OUTPUT ${full_lib_name_combined}
            COMMENT "Bundling ${lib_name_combined}"
            VERBATIM)
    else()
        message(FATAL_ERROR "Unknown bundle scenario!")
    endif()

    add_custom_target(bundling_target ALL DEPENDS ${full_lib_name_combined}) # NOTE: this line is actually the only one required to *just* build the bundled static library

    # Everything below here might one day be useful to link against the bundled library from CMake configurations, but I currently do not know how to accomplish that
    add_dependencies(bundling_target ${lib_name})

    add_library(${lib_name_combined} STATIC IMPORTED)
    set_target_properties(${lib_name_combined}
      PROPERTIES
        IMPORTED_LOCATION ${full_lib_name_combined}
        INTERFACE_INCLUDE_DIRECTORIES $<TARGET_PROPERTY:${lib_name},INTERFACE_INCLUDE_DIRECTORIES>)
    add_dependencies(${lib_name_combined} bundling_target)
endfunction()
#=======================================================================================================================
# END OF FUNCTION: bundle_static_library_simplified
#=======================================================================================================================
