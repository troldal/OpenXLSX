
@echo off

SET root=%cd%
SET zlib=lib\zlib-1.2.11
SET libzip=lib\libzip-1.8.0

if not exist "%zlib%" goto error_zlib_not_found
if not exist "%libzip%" goto error_libzip_not_found

:compile_zlib
if exist "%zlib%\build\.compiled" goto compile_libzip
echo =============================
echo Compiling zlib...
echo =============================
cd "%zlib%"
mkdir build
cd "build"
cmake ".." -DCMAKE_INSTALL_PREFIX="%root%/lib/install"
if %ERRORLEVEL% GEQ 1 goto error_zlib
cmake --build "." --config Debug --target install
if %ERRORLEVEL% GEQ 1 goto error_zlib
cmake --build "." --config Release --target install
if %ERRORLEVEL% GEQ 1 goto error_zlib
echo "OK" > ".compiled"
cd "..\..\.."

:compile_libzip
if exist "%libzip%\build\.compiled" goto compile_libzippp
echo =============================
echo Compiling libzip...
echo =============================
cd "%libzip%"
mkdir "build"
cd "build"
cmake ".." -DCMAKE_INSTALL_PREFIX="%root%/lib/install" -DCMAKE_PREFIX_PATH="%root%/lib/install" -DENABLE_COMMONCRYPTO=OFF -DENABLE_GNUTLS=OFF -DENABLE_MBEDTLS=OFF -DENABLE_OPENSSL=OFF -DENABLE_WINDOWS_CRYPTO=ON -DBUILD_TOOLS=OFF -DBUILD_REGRESS=OFF -DBUILD_EXAMPLES=OFF -DBUILD_DOC=OFF
if %ERRORLEVEL% GEQ 1 goto error_libzip
cmake --build . --config Debug --target install
if %ERRORLEVEL% GEQ 1 goto error_libzip
cmake --build . --config Release --target install
if %ERRORLEVEL% GEQ 1 goto error_libzip
echo "OK" > ".compiled"
cd "..\..\.."

:compile_libzippp
rmdir /q /s "dist"
mkdir "dist"
cd "dist"
mkdir "release"
copy "..\%zlib%\build\Release\zlib.dll" release
copy "..\%libzip%\build\lib\Release\zip.dll" release
copy "..\src\libzippp.h" release

mkdir "debug"
copy "..\%zlib%\build\Debug\zlibd.dll" debug
copy "..\%libzip%\build\lib\Debug\zip.dll" debug
copy "..\src\libzippp.h" debug
cd ".."

echo =============================
echo Compiling (shared) lizippp...
echo =============================
rmdir /q /s "build"
mkdir "build"
cd "build"
cmake .. -DCMAKE_PREFIX_PATH="%root%/lib/install" -DBUILD_SHARED_LIBS=ON
if %ERRORLEVEL% GEQ 1 goto error_libzippp
cd ".."
cmake --build build --config Debug
if %ERRORLEVEL% GEQ 1 goto error_libzippp
cmake --build build --config Release
if %ERRORLEVEL% GEQ 1 goto error_libzippp

copy "build\Release\libzippp_shared_test.exe" "dist/release"
copy "build\Release\libzippp.dll" "dist/release"
copy "build\Release\libzippp.lib" "dist/release"
copy "build\Debug\libzippp_shared_test.exe" "dist/debug"
copy "build\Debug\libzippp.dll" "dist/debug"
copy "build\Debug\libzippp.lib" "dist/debug"

echo =============================
echo Compiling (static) lizippp...
echo =============================
rmdir /q /s "build"
mkdir "build"
cd "build"
cmake .. -DCMAKE_PREFIX_PATH="%root%/lib/install" -DBUILD_SHARED_LIBS=OFF
if %ERRORLEVEL% GEQ 1 goto error_libzippp
cd ".."
cmake --build build --config Debug
if %ERRORLEVEL% GEQ 1 goto error_libzippp
cmake --build build --config Release
if %ERRORLEVEL% GEQ 1 goto error_libzippp

copy "build\Release\libzippp_static_test.exe" "dist/release"
copy "build\Release\libzippp_static.lib" "dist/release"
copy "build\Debug\libzippp_static_test.exe" "dist/debug"
copy "build\Debug\libzippp_static.lib" "dist/debug"

goto end

:error_zlib_not_found
echo [ERROR] The path was not found: %zlib%.
echo         You have to download zlib 1.2.11 and put in the folder %zlib%.
goto end

:error_zlib
echo [ERROR] Unable to compile zlib
goto end

:error_libzip_not_found
echo [ERROR] The path was not found: %libzip%.
echo         You have to download libzip 1.8.0 and put it in the folder %libzip%.
goto end

:error_libzip
echo [ERROR] Unable to compile libzip
goto end

:error_vs2012_not_found
echo [ERROR] VS2012 was not found (path not found: %vs2012devprompt%).
goto end

:error_libzippp
echo [ERROR] Unable to compile libzippp
goto end

:end
cd "%root%"
cmd
