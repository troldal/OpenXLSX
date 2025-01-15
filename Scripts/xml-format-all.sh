#!/bin/sh

SCRIPT_PATH=`dirname "$0"`
# echo "SCRIPT_PATH is $SCRIPT_PATH"

EXTENSIONS="xml rels"
EXTENSIONS_FOR_ECHO="\"xml\", \"rels\""
for arg in "$@" ; do
    if [ "$arg" != "xml" ] && [ "$arg" != "rels" ]; then
        EXTENSIONS="$EXTENSIONS $arg"
        EXTENSIONS_FOR_ECHO="$EXTENSIONS_FOR_ECHO, \"$arg\""
    else
        echo "ignoring default extension [\"xml\" or \"rels\"]"
    fi
done

echo "applying xmllint recursively to all files in folder ending in any of [$EXTENSIONS_FOR_ECHO]"
# echo "find . -type f -exec $SCRIPT_PATH/guarded-xml-format.sh {} $EXTENSIONS \;"
find . -type f -exec "$SCRIPT_PATH/guarded-xml-format.sh" {} $EXTENSIONS \;
