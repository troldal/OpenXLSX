# ============================================================================
# Author: Lars Uffmann (coding23@uffmann.name)
# Purpose: Patch an upstream package config (.pc) file Libs property to include an additional subdirectory {exec_prefix}/lib/SUBDIR

# Function: pkg_config_add_library_subdir
# Purpose: patch a package config file IN_PKG_CONFIG to contain an additional library directory LIB_SUBDIR and write the output to OUT_PKG_CONFIG
# Parameters:
#   IN_PKG_CONFIG
#   OUT_PKG_CONFIG
#   LIB_SUBDIR
# Returns: Writes the patched output file to the indicated location. Calling process has to verify success by checking existence of OUT_PKG_CONFIG
#
function(pkg_config_add_library_subdir IN_PKG_CONFIG OUT_PKG_CONFIG LIB_SUBDIR)
    message( DEBUG "prefixing ${IN_PKG_CONFIG} Libs: setting with '-L{exec_prefix}/lib/${LIB_SUBDIR}' and writing to ${OUT_PKG_CONFIG}" )

    set(NEW_LIB_DIR "-L\${exec_prefix}/lib/${LIB_SUBDIR}")

    file(READ "${IN_PKG_CONFIG}" PKG_CONFIG_RAW)

    # normalize newlines, then join continuation lines, if any
    string(REPLACE "\r\n" "\n"  PKG_CONFIG_RAW "${PKG_CONFIG_RAW}")
    string(REPLACE "\r" "\n"    PKG_CONFIG_RAW "${PKG_CONFIG_RAW}")
    string(REPLACE "\\\n" " "   PKG_CONFIG_RAW "${PKG_CONFIG_RAW}")

    # optional: collapse multiple whitespaces
    string(REGEX REPLACE "[ \t]+" " " PKG_CONFIG_RAW "${PKG_CONFIG_RAW}")

    # save and reload file as list of strings:
    file(WRITE "${OUT_PKG_CONFIG}" "${PKG_CONFIG_RAW}")
    file(STRINGS "${OUT_PKG_CONFIG}" PKG_CONFIG_LINES)

    # initialize text block variables
    set(BEFORE "")
    set(LIBS "")
    set(AFTER "")
    set(found_libs FALSE)

    # message( STATUS "PKG_CONFIG_LINES is:" )
    foreach(line IN LISTS PKG_CONFIG_LINES)
        # message( STATUS "   ${line}" )
        string(STRIP "${line}" line)
        if(line MATCHES "^[Ll]ibs[ \t]*:.*")
            string(REGEX REPLACE "^[Ll]ibs[ \t]*:[ \t]*" "" _Libs_Line "${line}")
            # message( STATUS "   This is the Libs line: |${_Libs_Line}|" )
            if(NOT found_libs)
                set(found_libs TRUE)
                set(LIBS "Libs: ${NEW_LIB_DIR}")    # begin Libs line with NEW_LIB_DIR
            endif()
            if(NOT "${_Libs_Line}" STREQUAL "") # for non-empty Libs entries
                set(LIBS "${LIBS} ${_Libs_Line}")   # append them to the collapsed Libs line
            endif()
        else()                                  # line is not a Libs line
            if(NOT found_libs)                      # if no Libs line was found so far
                set(BEFORE "${BEFORE}${line}\n")        # append line to "BEFORE" block
            else()                                  # else: at least one Libs line was already processed
                set(AFTER "${AFTER}${line}\n")          # append line to "AFTER" block
            endif()
        endif()
    endforeach()

    # if a LIBS line was generated, append a line break
    if(NOT "${LIBS}" STREQUAL "")
        set(LIBS "${LIBS}\n")
    endif()

    # add header comment explaining the automated patch and write the finalized outfile
    set(HEADER "# Patched by OpenXLSX installer: prepended ${NEW_LIB_DIR} to Libs to accommodate custom library location\n")
    file(WRITE "${OUT_PKG_CONFIG}" "${HEADER}${BEFORE}${LIBS}${AFTER}")
endfunction()
