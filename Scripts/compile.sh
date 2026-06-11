#!/bin/sh
if [ ! -e "$1" ]; then
    echo "must pass at least one input file as parameter!"
    exit 1
fi

SourceModules=
argCount=0
for arg in $@; do
    SourceModules="$SourceModules $lastArg"
    lastArg=$arg
    argCount=$((argCount+1))
done
if [ $argCount -gt 1 ]; then
    echo "using last argument '$lastArg' as output file name"
    OutputFile=$lastArg
else
    SourceModules=$lastArg
    OutputFile=a.out
fi

echo "g++ `pkg-config --cflags OpenXLSX` $SourceModules `pkg-config --static --libs OpenXLSX` -o \"$OutputFile\""
g++ `pkg-config --cflags OpenXLSX` $SourceModules `pkg-config --static --libs OpenXLSX` -o "$OutputFile"
