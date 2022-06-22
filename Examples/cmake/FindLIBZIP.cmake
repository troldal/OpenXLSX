find_package(ZLIB REQUIRED)

find_path(LIBZIP_INCLUDE_DIR NAMES zip.h)
mark_as_advanced(LIBZIP_INCLUDE_DIR)

find_library(LIBZIP_LIBRARY NAMES zip)
mark_as_advanced(LIBZIP_LIBRARY)

get_filename_component(_libzip_libdir ${LIBZIP_LIBRARY} DIRECTORY)
find_file(_libzip_pkgcfg libzip.pc
    HINTS ${_libzip_libdir} ${LIBZIP_INCLUDE_DIR}/..
    PATH_SUFFIXES pkgconfig lib/pkgconfig
    NO_DEFAULT_PATH
)

include(FindPackageHandleStandardArgs)
find_package_handle_standard_args(
    LIBZIP 
    REQUIRED_VARS
        LIBZIP_LIBRARY
        LIBZIP_INCLUDE_DIR
        _libzip_pkgcfg
)

if (LIBZIP_FOUND)
    if (NOT TARGET libzip::zip)
        add_library(libzip::zip UNKNOWN IMPORTED)
        set_target_properties(libzip::zip
            PROPERTIES
                INTERFACE_INCLUDE_DIRECTORIES ${LIBZIP_INCLUDE_DIR}
                INTERFACE_LINK_LIBRARIES ZLIB::ZLIB
                IMPORTED_LOCATION ${LIBZIP_LIBRARY}
        )
        # (Ab)use the (always) installed pkgconfig file to check if BZip2 is required
        file(STRINGS ${_libzip_pkgcfg} _have_extra_libs REGEX Libs)
        if(_have_extra_libs MATCHES "-lbz2")
            find_package(BZip2 REQUIRED)
            set_property(TARGET libzip::zip APPEND PROPERTY INTERFACE_LINK_LIBRARIES BZip2::BZip2)
        endif()
        if(_have_extra_libs MATCHES "-lcrypto")
            find_package(OpenSSL REQUIRED)
            set_property(TARGET libzip::zip APPEND PROPERTY INTERFACE_LINK_LIBRARIES OpenSSL::Crypto)
        endif()
        if(_have_extra_libs MATCHES "-lgnutls")
            find_package(GnuTLS::GnuTLS REQUIRED)
            set_property(TARGET libzip::zip APPEND PROPERTY INTERFACE_LINK_LIBRARIES GnuTLS::GnuTLS)
        endif()
        if(_have_extra_libs MATCHES "lgnutls")
            find_package(Nettle::Nettle REQUIRED)
            set_property(TARGET libzip::zip APPEND PROPERTY INTERFACE_LINK_LIBRARIES Nettle::Nettle)
        endif()
        if(_have_extra_libs MATCHES "-llzma")
            find_package(LibLZMA REQUIRED)
            set_property(TARGET libzip::zip APPEND PROPERTY INTERFACE_LINK_LIBRARIES LibLZMA::LibLZMA)
        endif()
        if(_have_extra_libs MATCHES "-lz")
            find_package(ZLIB REQUIRED)
            set_property(TARGET libzip::zip APPEND PROPERTY INTERFACE_LINK_LIBRARIES ZLIB::ZLIB)
        endif()
    endif()
endif()
