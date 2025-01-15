#!/bin/sh

# ===== BEGIN Detect script path and store in SCRIPT_PATH =====
ELIMINATE_SLASHLESS=`echo "$0" | sed 's|[^/]*$||'`
# echo "ELIMINATE_SLASHLESS is $ELIMINATE_SLASHLESS"
if [ "$ELIMINATE_SLASHLESS" = "" ]; then
    SCRIPT_PATH="."
else
    SCRIPT_PATH=`echo "$0" | sed 's|/[^/]*$||'`
fi

# echo "SCRIPT_PATH is $SCRIPT_PATH"
# ===== END Detect script path and store in SCRIPT_PATH =====


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

exit 0
