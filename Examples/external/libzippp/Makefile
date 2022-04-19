CFLAGS=-W -Wall -Wextra -ansi -pedantic -std=c++11
OBJ=obj
LIB=lib
ZLIB_VERSION=1.2.11
ZLIB=$(LIB)/zlib-$(ZLIB_VERSION)
LIBZIP_VERSION=1.8.0
LIBZIP=$(LIB)/libzip-$(LIBZIP_VERSION)
LIBZIP_CMAKE=-DENABLE_COMMONCRYPTO=OFF -DENABLE_GNUTLS=OFF -DENABLE_MBEDTLS=OFF 
CRYPTO_FLAGS=-lssl -lcrypto

# for optimal compilation speed, should be <nb_proc>+1
NBPROC=5

all: libzippp-static libzippp-shared

libzippp-compile:
	rm -rf $(OBJ)
	mkdir $(OBJ)
	$(CXX) -g -fPIC -c -I$(LIBZIP)/lib -I$(LIBZIP)/build -o $(OBJ)/libzippp.o $(CFLAGS) src/libzippp.cpp

libzippp-static: libzippp-compile
	ar rvs libzippp.a $(OBJ)/libzippp.o

libzippp-shared: libzippp-compile
	$(CXX) -shared -o libzippp.so $(OBJ)/libzippp.o

libzippp-tests: libzippp-static libzippp-shared
	if [ -d $(ZLIB) ]; then \
		$(CXX) -o test_static -g -I$(ZLIB) -I$(LIBZIP)/lib -I$(LIBZIP)/build -Isrc $(CFLAGS) tests/tests.cpp libzippp.a $(LIBZIP)/build/lib/libzip.a $(ZLIB)/libz.a -lbz2 -llzma $(CRYPTO_FLAGS); \
		$(CXX) -o test_shared -g -I$(ZLIB) -I$(LIBZIP)/lib -I$(LIBZIP)/build -Isrc $(CFLAGS) tests/tests.cpp -L. -L$(LIBZIP)/build/lib -L$(ZLIB) -lzippp -lzip -lz -lbz2 -llzma $(CRYPTO_FLAGS) -Wl,-rpath=.; \
	else \
		$(CXX) -o test_static -g -I$(LIBZIP)/lib -Isrc $(CFLAGS) tests/tests.cpp libzippp.a $(LIBZIP)/build/lib/libzip.a -lz -lbz2 -llzma $(CRYPTO_FLAGS); \
		$(CXX) -o test_shared -g -I$(LIBZIP)/lib -Isrc $(CFLAGS) tests/tests.cpp -L. -L$(LIBZIP)/build/lib -lzippp -lzip -lz -lbz2 -llzma $(CRYPTO_FLAGS) -Wl,-rpath=.; \
	fi;

clean-tests:
	@rm -rf *.zip

tests: libzippp-tests clean-tests
	LD_LIBRARY_PATH="$(LIBZIP)/build/lib" valgrind --suppressions=ld.supp --leak-check=full ./test_shared
	valgrind --suppressions=ld.supp --leak-check=full ./test_static

clean:
	@rm -rf libzippp.a libzippp.so
	@rm -rf $(OBJ)
	@rm -rf test_shared test_static
	@rm -f windows-ready_to_compile.zip

mrproper: clean
	@rm -rf $(LIBZIP)
	@rm -rf $(LIB)/libzip-$(LIBZIP_VERSION)
	@rm -rf $(LIB)/zlib-$(ZLIB_VERSION)
	@rm -rf $(LIB)/*.tar.gz
	
# ZLIB targets
	
zlib-init:
	mkdir -p $(LIB)
	
zlib-download: zlib-init
	wget -c -O "$(ZLIB).tar.gz" "https://github.com/madler/zlib/archive/v$(ZLIB_VERSION).tar.gz"

zlib-unzip: zlib-download
	cd $(LIB) && tar -xf zlib-$(ZLIB_VERSION).tar.gz

zlib-configure: zlib-unzip
	cd $(ZLIB) && ./configure

zlib-compile: zlib-configure
	cd $(ZLIB) && make -j$(NBPROC)

zlib: zlib-compile

# LIZIP targets

libzip-init:
	mkdir -p $(LIB)

libzip-download: libzip-init
	wget -c -O "$(LIBZIP).tar.gz" "https://www.nih.at/libzip/libzip-$(LIBZIP_VERSION).tar.gz"

libzip-unzip: libzip-download
	cd $(LIB) && tar -xf libzip-$(LIBZIP_VERSION).tar.gz

libzip-patch: libzip-unzip
	if [ -f $(LIB)/libzip-$(LIBZIP_VERSION)-windows.patch ]; then \
		cd $(LIB)/libzip-$(LIBZIP_VERSION) && patch -p1 < ../libzip-$(LIBZIP_VERSION)-windows.patch; \
	fi;

libzip-build-folder:
	mkdir -p $(LIBZIP)/build;

libzip-build-shared: libzip-patch libzip-build-folder
	if [ -d "$(ZLIB)" ]; then \
		cd $(LIBZIP)/build && cmake .. -DZLIB_LIBRARY_RELEASE=../../../$(ZLIB)/libz.so -DZLIB_INCLUDE_DIR=../../../$(ZLIB) -DBUILD_SHARED_LIBS=ON $(LIBZIP_CMAKE) && make -j$(NBPROC);  \
	else \
		cd $(LIBZIP)/build && cmake .. -DBUILD_SHARED_LIBS=ON $(LIBZIP_CMAKE) && make -j$(NBPROC);  \
	fi;
	
libzip-build-static: libzip-patch libzip-build-folder
	if [ -d "$(ZLIB)" ]; then \
		cd $(LIBZIP)/build && cmake .. -DZLIB_LIBRARY_RELEASE=../../../$(ZLIB)/libz.a -DZLIB_INCLUDE_DIR=../../../$(ZLIB) -DBUILD_SHARED_LIBS=OFF $(LIBZIP_CMAKE) && make -j$(NBPROC);  \
	else \
		cd $(LIBZIP)/build && cmake .. -DBUILD_SHARED_LIBS=OFF $(LIBZIP_CMAKE) && make -j$(NBPROC);  \
	fi;

libzip: libzip-build-shared libzip-build-static

# LIBRARIES TARGET

libraries: zlib libzip

libraries-download: zlib-download libzip-download

# Windows compilation preparation

windows: mrproper zlib-download zlib-unzip libzip-download libzip-unzip libzip-patch
	zip -r windows-ready_to_compile.zip lib src tests cmake compile.bat CMakeLists.txt Config.cmake.in
