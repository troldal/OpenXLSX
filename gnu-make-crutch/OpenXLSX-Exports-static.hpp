
#ifndef OPENXLSX_EXPORT_H
#define OPENXLSX_EXPORT_H

#ifdef OPENXLSX_STATIC_DEFINE
#  define OPENXLSX_EXPORT
#  define OPENXLSX_HIDDEN
#else
#  ifndef OPENXLSX_EXPORT
#    ifdef OpenXLSX_EXPORTS
        /* We are building this library */
#      define OPENXLSX_EXPORT 
#    else
        /* We are using this library */
#      define OPENXLSX_EXPORT 
#    endif
#  endif

#  ifndef OPENXLSX_HIDDEN
#    define OPENXLSX_HIDDEN 
#  endif
#endif

#ifndef OPENXLSX_DEPRECATED
#  define OPENXLSX_DEPRECATED __attribute__ ((__deprecated__))
#endif

#ifndef OPENXLSX_DEPRECATED_EXPORT
#  define OPENXLSX_DEPRECATED_EXPORT OPENXLSX_EXPORT OPENXLSX_DEPRECATED
#endif

#ifndef OPENXLSX_DEPRECATED_NO_EXPORT
#  define OPENXLSX_DEPRECATED_NO_EXPORT OPENXLSX_HIDDEN OPENXLSX_DEPRECATED
#endif

#if 0 /* DEFINE_NO_DEPRECATED */
#  ifndef OPENXLSX_NO_DEPRECATED
#    define OPENXLSX_NO_DEPRECATED
#  endif
#endif

#endif /* OPENXLSX_EXPORT_H */
