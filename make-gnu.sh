#!/bin/sh

# assemble all arguments with individual quote pairs to echo correctly with string preservation
ARGS=
for arg in "$@"
do
    ARGS="$ARGS \"$arg\""
done

echo "make --makefile Makefile.GNU $ARGS"
make --makefile Makefile.GNU "$@"
