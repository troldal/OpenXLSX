# GNU Makefile for OpenXLSX modules & demos
# Version: 2024-09-29 13:45 CEST

# NOTE: lib and library are available as aliases for OpenXLSX, to build the static library file
TARGETS=OpenXLSX Demo1 Demo2 Demo3 Demo4 Demo5 Demo6 Demo7 Demo8 Demo9 Demo10

DEMOS_SRC_DIR=Examples
BIN_DIR=output

CC=gcc
CXX=g++

# /usr/bin/ar: create archive without symbol table
AR=ar
# /usr/bin/ranlib: add symbol table to archive created with ar
RANLIB=ranlib


OPTIMIZATION_FLAGS=
#OPTIMIZATION_FLAGS=-O3

# === BEGIN: Detect OS platform based on gcc compiler defines ===
# Actual detection command:
DETECT_CMD="echo | $(CC) -dM -E -"

# Test commands to "simulate" a platform
# DETECT_CMD="echo \"\#define WIN32 1\""
# DETECT_CMD="echo \"\#define _WIN32 1\""
# DETECT_CMD="echo \"\#define __WIN32__ 1\""
# DETECT_CMD="echo \"\#define _WIN64 1\""
# DETECT_CMD="printf '\#define WIN32 1\n\#define _WIN32 1\n\#define __WIN32__ 1\n\#define _WIN64 1'" # test all defines together with line breaks
# DETECT_CMD="echo \"\#define __linux__ 1\""
# DETECT_CMD="echo \"\#define __CYGWIN__ 1\""

__linux__ := $(shell "$(DETECT_CMD)" | grep "\b__linux__\b")
ifdef __linux__
	Linux=yes
else
	Linux=no
endif

WIN32 := $(shell "$(DETECT_CMD)" | grep "\bWIN32\b")
_WIN32 := $(shell "$(DETECT_CMD)" | grep "\b_WIN32\b")
__WIN32__ := $(shell "$(DETECT_CMD)" | grep "\b__WIN32__\b")
_WIN64 := $(shell "$(DETECT_CMD)" | grep "\b_WIN64\b")
ifdef WIN32
	Windows=yes
else ifdef _WIN32
	Windows=yes
else ifdef __WIN32__
	Windows=yes
else ifdef _WIN64
	Windows=yes
else
	Windows=no
endif

__CYGWIN__ := $(shell "$(DETECT_CMD)" | grep "\b__CYGWIN__\b")
ifdef __CYGWIN__
	Cygwin=yes
else
	Cygwin=no
endif
# === END: detect OS platform based on gcc compiler defines ===

OPENXLSX_DIR=OpenXLSX
SRC_DIR=sources
INCLUDE_DIR=headers
OBJ_DIR=obj
STATIC_LIBRARY=libOpenXLSX.a

USE_NOWIDE=no
ifeq ("$(Windows)", "yes")
	# enable nowide on windows. CAUTION: do not append this comment to the variable assignment, it will break the ifeq check below
	USE_NOWIDE=yes
endif

# SHARED_SUBDIR=shared
EXTERNAL_SUBDIR=external
NOWIDE_SUBDIR=$(EXTERNAL_SUBDIR)/nowide
PUGIXML_SUBDIR=$(EXTERNAL_SUBDIR)/pugixml
ZIPPY_SUBDIR=$(EXTERNAL_SUBDIR)/zippy

# library / utility objects
# OBJS_LICENSE=license.o
# OBJS_SHARED=$(OBJS_LICENSE)
OBJS_PUGIXML= # used as header-only module
OBJS_ZIPPY=   # header-only module
OBJS_OPENXLSX=XLCell.o XLCellIterator.o XLCellRange.o XLCellReference.o XLCellValue.o XLColor.o XLColumn.o XLComments.o XLContentTypes.o XLDateTime.o XLDocument.o XLDrawing.o XLFormula.o XLMergeCells.o XLProperties.o XLRelationships.o XLRow.o XLRowData.o XLSharedStrings.o XLSheet.o XLStyles.o XLTables.o XLWorkbook.o XLXmlData.o XLXmlFile.o XLXmlParser.o XLZipArchive.o

# create a version of OBJS_OPENXLSX that already has the correct prefix so that it can be used for linking without further modification
OBJS_OPENXLSX_PREFIXED=$(addprefix $(OBJ_DIR)/$(OPENXLSX_DIR)/,$(OBJS_OPENXLSX))

# OBJS_DEMOS=$(OBJS_OPENXLSX_PREFIXED) # no longer needed - using static library instead
OBJS_DEMOS=$(BIN_DIR)/$(STATIC_LIBRARY) # use static library for linking demos

PROJECT_FLAGS=
ifeq ("$(USE_NOWIDE)", "yes")
	PROJECT_FLAGS=-DENABLE_NOWIDE
endif


# additional include directories for 1) OPENXLSX project headers and 2) external nowide (parent) folder
ADDITIONAL_INCLUDE_FLAGS=-I$(OPENXLSX_DIR) -I$(OPENXLSX_DIR)/$(EXTERNAL_SUBDIR)

# application objects

# Demo0 is a placeholder for test programs - if used, has to be explicitly provided as build target (not included in defaults)
OBJS_DEMO0=Demo0.o
OBJS_OPENXLSX_DEMO0=$(OBJS_DEMOS)
LDLIBS_DEMO0=

OBJS_DEMO1=Demo1.o
OBJS_OPENXLSX_DEMO1=$(OBJS_DEMOS)
LDLIBS_DEMO1=

OBJS_DEMO2=Demo2.o
OBJS_OPENXLSX_DEMO2=$(OBJS_DEMOS)
LDLIBS_DEMO2=

OBJS_DEMO3=Demo3.o
OBJS_OPENXLSX_DEMO3=$(OBJS_DEMOS)
LDLIBS_DEMO3=

OBJS_DEMO4=Demo4.o
OBJS_OPENXLSX_DEMO4=$(OBJS_DEMOS)
LDLIBS_DEMO4=

OBJS_DEMO5=Demo5.o
OBJS_OPENXLSX_DEMO5=$(OBJS_DEMOS)
LDLIBS_DEMO5=

OBJS_DEMO6=Demo6.o
OBJS_OPENXLSX_DEMO6=$(OBJS_DEMOS)
LDLIBS_DEMO6=

OBJS_DEMO7=Demo7.o
OBJS_OPENXLSX_DEMO7=$(OBJS_DEMOS)
LDLIBS_DEMO7=

OBJS_DEMO8=Demo8.o
OBJS_OPENXLSX_DEMO8=$(OBJS_DEMOS)
LDLIBS_DEMO8=

OBJS_DEMO9=Demo9.o
OBJS_OPENXLSX_DEMO9=$(OBJS_DEMOS)
LDLIBS_DEMO9=

OBJS_DEMO10=Demo10.o
OBJS_OPENXLSX_DEMO10=$(OBJS_DEMOS)
LDLIBS_DEMO10=

SANITIZE_FLAGS=
# SANITIZE_FLAGS=-fsanitize=address
# SANITIZE_FLAGS=-fsanitize=address -fsanitize=leak
SANITIZE_LIBS=
# SANITIZE_LIBS=-llsan

# 2024-08-18 warning state:
#   misleading-indentation: 0 warnings
#   sign-compare: 0 warnings
#   unknown-pragmas: 0 (disabled warning about previous 144 lines upsetting gcc/g++ by pragma)
#   unused-function: 3 (disabled by pragma)
#     OpenXLSX/sources/XLProperties.cpp: std::vector<std::string> headingPairsCategoriesStrings(XMLNode docNode)
#     OpenXLSX/sources/XLDocument.cpp: bool fileExists(const std::string& fileName)
#                                      bool isDirectory(const std::string& fileName)
ENABLED_WARNING_FLAGS=-Wformat -Wformat-signedness -Wall -Wpedantic -Wextra
DISABLED_WARNING_FLAGS=

GENERIC_FLAGS=$(ADDITIONAL_INCLUDE_FLAGS) $(PROJECT_FLAGS) -fno-common $(SANITIZE_FLAGS) $(OPTIMIZATION_FLAGS) $(ENABLED_WARNING_FLAGS) $(DISABLED_WARNING_FLAGS)

CPPFLAGS=$(GENERIC_FLAGS) -std=c++17
# CPPFLAGS=$(GENERIC_FLAGS) -std=c++17 -D_GNU_SOURCE -D_POSIX_C_SOURCE=200809L -D_XOPEN_SOURCE=700

LDFLAGS=$(SANITIZE_FLAGS)
# LDFLAGS=-fno-common

# TEST: truncate binary to actually used code (https://gcc.gnu.org/onlinedocs/gnat_ugn/Compilation-options.html)
# GENERIC_FLAGS+= -ffunction-sections -fdata-sections
# LDFLAGS+= -Wl,--gc-sections


# precompiled libs go here
LDLIBS=$(SANITIZE_LIBS)
# LDLIBS=-lrt -pthread -lboost_program_options $(SANITIZE_LIBS) # example to add libraries if needed


## additional tools for text substitution:
# call tools like so:
# VAR = MixedCaseText
# LOWER_VAR = $(call lc,$(VAR))
#   create all lowercase variable:
# lc = $(subst A,a,$(subst B,b,$(subst C,c,$(subst D,d,$(subst E,e,$(subst F,f,$(subst G,g,$(subst H,h,$(subst I,i,$(subst J,j,$(subst K,k,$(subst L,l,$(subst M,m,$(subst N,n,$(subst O,o,$(subst P,p,$(subst Q,q,$(subst R,r,$(subst S,s,$(subst T,t,$(subst U,u,$(subst V,v,$(subst W,w,$(subst X,x,$(subst Y,y,$(subst Z,z,$1))))))))))))))))))))))))))
#   create all uppercase variable:
uc = $(subst a,A,$(subst b,B,$(subst c,C,$(subst d,D,$(subst e,E,$(subst f,F,$(subst g,G,$(subst h,H,$(subst i,I,$(subst j,J,$(subst k,K,$(subst l,L,$(subst m,M,$(subst n,N,$(subst o,O,$(subst p,P,$(subst q,Q,$(subst r,R,$(subst s,S,$(subst t,T,$(subst u,U,$(subst v,V,$(subst w,W,$(subst x,X,$(subst y,Y,$(subst z,Z,$1))))))))))))))))))))))))))


## gnu make override to define the default make directive, only takes a single target
.DEFAULT_GOAL := default
default: $(TARGETS)

all: $(TARGETS)

# the $(NOOP) rules are only needed for the dependency / forward to $(BIN_DIR)
Demo0: $(BIN_DIR)/Demo0
	$(NOOP)
Demo1: $(BIN_DIR)/Demo1
	$(NOOP)
Demo2: $(BIN_DIR)/Demo2
	$(NOOP)
Demo3: $(BIN_DIR)/Demo3
	$(NOOP)
Demo4: $(BIN_DIR)/Demo4
	$(NOOP)
Demo5: $(BIN_DIR)/Demo5
	$(NOOP)
Demo6: $(BIN_DIR)/Demo6
	$(NOOP)
Demo7: $(BIN_DIR)/Demo7
	$(NOOP)
Demo8: $(BIN_DIR)/Demo8
	$(NOOP)
Demo9: $(BIN_DIR)/Demo9
	$(NOOP)
Demo10: $(BIN_DIR)/Demo10
	$(NOOP)

# re-direct from aliases
lib: OpenXLSX
library: OpenXLSX

# static library target
OpenXLSX: $(BIN_DIR)/$(STATIC_LIBRARY)

# delete existing file, create static library archive file & add symbol table to archive
# NOTE: if existing library file is not explicitly deleted, ar will fail to update a modified object file
$(BIN_DIR)/$(STATIC_LIBRARY): $(OBJS_OPENXLSX_PREFIXED) | $(BIN_DIR)
	rm -f $@ # explicitly delete existing library file
	$(AR) qc $@ $(addprefix $(OBJ_DIR)/$(OPENXLSX_DIR)/,$(OBJS_OPENXLSX))
	$(RANLIB) $@

# complex rule for BIN_DIR targets: include objects from OBJS_<target> and libraries from LDLIBS_<target>
.SECONDEXPANSION:
$(BIN_DIR)/%: $$(addprefix $(OBJ_DIR)/,$$(OBJS_$$(call uc,$$*))) $$(OBJS_OPENXLSX_$$(call uc,$$*)) | $(BIN_DIR)
	$(CXX) $(LDFLAGS) $^ $(LDLIBS) $(LDLIBS_$(call uc,$*)) -o $@

# rule for .cpp files in DEMOS_SRC_DIR
$(OBJ_DIR)/%.o: $(addprefix $(DEMOS_SRC_DIR)/,%.cpp) | $(OBJ_DIR) $(OPENXLSX_DIR)/OpenXLSX-Exports.hpp
	$(CXX) $(CPPFLAGS) -I$(OPENXLSX_DIR)/$(INCLUDE_DIR) -I$(OPENXLSX_DIR)/$(NOWIDE_SUBDIR) -c $< -o $@

# rule for OpenXLSX .cpp files
$(OBJ_DIR)/$(OPENXLSX_DIR)/%.o: $(OPENXLSX_DIR)/$(SRC_DIR)/%.cpp | $(OBJ_DIR) $(OBJ_DIR)/$(OPENXLSX_DIR)/ $(OPENXLSX_DIR)/OpenXLSX-Exports.hpp
	$(CXX) $(CPPFLAGS) -I$(OPENXLSX_DIR)/$(INCLUDE_DIR) -I$(OPENXLSX_DIR)/$(NOWIDE_SUBDIR) -I$(OPENXLSX_DIR)/$(PUGIXML_SUBDIR) -I$(OPENXLSX_DIR)/$(ZIPPY_SUBDIR) -c $< -o $@

# rule for nowide .cpp files: N/A, as nowide is header-only

# rule for pugixml .cpp files: N/A, as pugixml is used as header only
$(OBJ_DIR)/$(PUGIXML_SUBDIR)/%.o: $(OPENXLSX_DIR)/$(PUGIXML_SUBDIR)/%.cpp | $(OBJ_DIR) $(OBJ_DIR)/$(PUGIXML_SUBDIR)/
	$(CXX) $(CPPFLAGS) -I$(OPENXLSX_DIR)/$(PUGIXML_SUBDIR) -c $< -o $@

# rule for zippy .cpp files: N/A, as zippy is header only

$(OPENXLSX_DIR)/OpenXLSX-Exports.hpp:
	cp gnu-make-crutch/OpenXLSX-Exports.hpp $@
	
# # rule for shared .cpp files
# $(OBJ_DIR)/$(SHARED_SUBDIR)/%.o: $(OPENXLSX_DIR)/$(SRC_DIR)/$(SHARED_SUBDIR)/%.cpp | $(OBJ_DIR) $(OBJ_DIR)/$(SHARED_SUBDIR)/
# 	$(CXX) $(CPPFLAGS) -I$(INCLUDE_DIR)/$(SHARED_SUBDIR) -c $< -o $@


# create BIN_DIR if not existing
$(BIN_DIR):
	mkdir $@

# create OBJ_DIR if not existing
$(OBJ_DIR):
	mkdir $@

# $(OBJ_DIR)/$(SHARED_SUBDIR)/:
# 	mkdir $@

$(OBJ_DIR)/$(OPENXLSX_DIR)/:
	mkdir $@

$(OBJ_DIR)/$(PUGIXML_SUBDIR)/:
	mkdir $@

$(OBJ_DIR)/$(ZIPPY_SUBDIR)/:
	mkdir $@

# .SECONDARY with no prerequisites causes all targets to be treated
# as secondary (i.e., no target is removed because it is considered
# intermediate).
.SECONDARY:

# indicate all rules without target output
.PHONY: default all clean cleanObjects cleanTargets echo lib library OpenXLSX

clean: cleanObjects cleanTargets

cleanObjects:
	rm -f $(addprefix $(OBJ_DIR)/,$(OBJS_DEMO0) $(OBJS_DEMO1) $(OBJS_DEMO2) $(OBJS_DEMO3) $(OBJS_DEMO4) $(OBJS_DEMO5) $(OBJS_DEMO6) $(OBJS_DEMO7) $(OBJS_DEMO8) $(OBJS_DEMO9) $(OBJS_DEMO10)) \
	      $(BIN_DIR)/$(STATIC_LIBRARY) \
	      $(OBJS_OPENXLSX_PREFIXED) \
	      $(addprefix $(OBJ_DIR)/$(PUGIXML_SUBDIR)/,$(OBJS_PUGIXML)) \
	      $(addprefix $(OBJ_DIR)/$(ZIPPY_SUBDIR)/,$(OBJS_ZIPPY))
# 	      $(addprefix $(OBJ_DIR)/$(SHARED_SUBDIR)/,$(OBJS_SHARED))

cleanTargets:
	rm -f $(addprefix $(BIN_DIR)/,Demo0) $(addprefix $(BIN_DIR)/,$(TARGETS))


# have a look at the established configuration
echo:
	@echo "OpenXLSX Makefile"
	@echo "----------------------------------------------"
	@echo "TARGETS: $(TARGETS)"
	@echo "DEMOS_SRC_DIR: $(DEMOS_SRC_DIR)"
	@echo "BIN_DIR: $(BIN_DIR)"
	@echo
	@echo "CC    : $(CC)"
	@echo "CXX   : $(CXX)"
	@echo "AR    : $(AR)"
	@echo "RANLIB: $(RANLIB)"
	@echo
	@echo "OPTIMIZATION_FLAGS: $(OPTIMIZATION_FLAGS)"
	@echo
	@echo "Platform detection variables:"
	@echo "-----------------------------"
	@echo 'DETECT_CMD: $(DETECT_CMD)'
	@echo
	@echo "__linux__ : $(__linux__)"
	@echo "WIN32     : $(WIN32)"
	@echo "_WIN32    : $(_WIN32)"
	@echo "__WIN32__ : $(__WIN32__)"
	@echo "_WIN64    : $(_WIN64)"
	@echo "__CYGWIN__: $(__CYGWIN__)"
	@echo
	@echo "Platform detection results:"
	@echo "---------------------------"
	@echo "Linux  : $(Linux)"
	@echo "Windows: $(Windows)"
	@echo "Cygwin : $(Cygwin)"
	@echo
	@echo "OpenXLSX project configuration:"
	@echo "-------------------------------"
	@echo "OPENXLSX_DIR: $(OPENXLSX_DIR)"
	@echo "SRC_DIR: $(SRC_DIR)"
	@echo "INCLUDE_DIR: $(INCLUDE_DIR)"
	@echo "OBJ_DIR: $(OBJ_DIR)"
	@echo "STATIC_LIBRARY: $(STATIC_LIBRARY)"
	@echo
	@echo "USE_NOWIDE: $(USE_NOWIDE)"
	@echo
# 	@echo "SHARED_SUBDIR: $(SHARED_SUBDIR)"
	@echo "EXTERNAL_SUBDIR: $(EXTERNAL_SUBDIR)"
	@echo "PUGIXML_SUBDIR: $(PUGIXML_SUBDIR)"
	@echo "ZIPPY_SUBDIR: $(ZIPPY_SUBDIR)"
	@echo
# 	@echo "OBJS_LICENSE: $(OBJS_LICENSE)"
# 	@echo "OBJS_SHARED: $(OBJS_SHARED)"
	@echo "OBJS_PUGIXML: $(OBJS_PUGIXML)"
	@echo "OBJS_ZIPPY: $(OBJS_ZIPPY)"
	@echo "OBJS_OPENXLSX: $(OBJS_OPENXLSX)"
	@echo "OBJS_OPENXLSX_PREFIXED: $(OBJS_OPENXLSX_PREFIXED)"
	@echo "OBJS_DEMOS: $(OBJS_DEMOS)"
	@echo
	@echo "PROJECT_FLAGS: $(PROJECT_FLAGS)"
	@echo "ADDITIONAL_INCLUDE_FLAGS: $(ADDITIONAL_INCLUDE_FLAGS)"
	@echo
	@echo "OBJS_DEMO0: $(OBJS_DEMO0)"
	@echo "OBJS_OPENXLSX_DEMO0: $(OBJS_OPENXLSX_DEMO0)"
	@echo "LDLIBS_DEMO0: $(LDLIBS_DEMO0)"
	@echo "OBJS_DEMO1: $(OBJS_DEMO1)"
	@echo "OBJS_OPENXLSX_DEMO1: $(OBJS_OPENXLSX_DEMO1)"
	@echo "LDLIBS_DEMO1: $(LDLIBS_DEMO1)"
	@echo "OBJS_DEMO2: $(OBJS_DEMO2)"
	@echo "OBJS_OPENXLSX_DEMO2: $(OBJS_OPENXLSX_DEMO2)"
	@echo "LDLIBS_DEMO2: $(LDLIBS_DEMO2)"
	@echo "OBJS_DEMO3: $(OBJS_DEMO3)"
	@echo "OBJS_OPENXLSX_DEMO3: $(OBJS_OPENXLSX_DEMO3)"
	@echo "LDLIBS_DEMO3: $(LDLIBS_DEMO3)"
	@echo "OBJS_DEMO4: $(OBJS_DEMO4)"
	@echo "OBJS_OPENXLSX_DEMO4: $(OBJS_OPENXLSX_DEMO4)"
	@echo "LDLIBS_DEMO4: $(LDLIBS_DEMO4)"
	@echo "OBJS_DEMO5: $(OBJS_DEMO5)"
	@echo "OBJS_OPENXLSX_DEMO5: $(OBJS_OPENXLSX_DEMO5)"
	@echo "LDLIBS_DEMO5: $(LDLIBS_DEMO5)"
	@echo "OBJS_DEMO6: $(OBJS_DEMO6)"
	@echo "OBJS_OPENXLSX_DEMO6: $(OBJS_OPENXLSX_DEMO6)"
	@echo "LDLIBS_DEMO6: $(LDLIBS_DEMO6)"
	@echo "OBJS_DEMO7: $(OBJS_DEMO7)"
	@echo "OBJS_OPENXLSX_DEMO7: $(OBJS_OPENXLSX_DEMO7)"
	@echo "LDLIBS_DEMO7: $(LDLIBS_DEMO7)"
	@echo "OBJS_DEMO8: $(OBJS_DEMO8)"
	@echo "OBJS_OPENXLSX_DEMO8: $(OBJS_OPENXLSX_DEMO8)"
	@echo "LDLIBS_DEMO8: $(LDLIBS_DEMO8)"
	@echo "OBJS_DEMO9: $(OBJS_DEMO9)"
	@echo "OBJS_OPENXLSX_DEMO9: $(OBJS_OPENXLSX_DEMO9)"
	@echo "LDLIBS_DEMO9: $(LDLIBS_DEMO9)"
	@echo "OBJS_DEMO10: $(OBJS_DEMO10)"
	@echo "OBJS_OPENXLSX_DEMO10: $(OBJS_OPENXLSX_DEMO10)"
	@echo "LDLIBS_DEMO10: $(LDLIBS_DEMO10)"
	@echo
	@echo "SANITIZE_FLAGS: $(SANITIZE_FLAGS)"
	@echo "SANITIZE_LIBS: $(SANITIZE_LIBS)"
	@echo "DISABLED_WARNING_FLAGS: $(DISABLED_WARNING_FLAGS)"
	@echo "GENERIC_FLAGS: $(GENERIC_FLAGS)"
	@echo "CPPFLAGS: $(CPPFLAGS)"
	@echo
	@echo "LDFLAGS: $(LDFLAGS)"
	@echo "LDLIBS: $(LDLIBS)"
