#!/bin/sh

if [ ! -e $1 ]; then
	echo "provided argument \"$1\" does not exist"
	exit 1
fi

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


EXTENSIONS=""
ALLOWED_EXTENSION="false"

FILE=""
extension="(uninitialized)"

for arg in "$@" ; do
    if [ "$FILE" = "" ]; then
        FILE=$arg
        extension=`echo $FILE | sed -r 's/[^.]*(\.([^.]*))*$/\2/g'`
#         echo "for file $FILE detected extension: $extension"
    else
#         echo "allowed extension: $arg"
        if [ "$EXTENSIONS" != "" ]; then
            EXTENSIONS="$EXTENSIONS, "
        fi
        EXTENSIONS="$EXTENSIONS\"$arg\""
        if [ "$extension" = "$arg" ]; then
#             echo "file name matches allowed extension $arg"
            ALLOWED_EXTENSION="true"
break;
        fi
    fi
done

if [ "$ALLOWED_EXTENSION" = "true" ]; then
#     echo "found an allowed extension - if it safe to invoke $SCRIPT_PATH/xml-format.sh for $FILE"
    "$SCRIPT_PATH/xml-format.sh" $FILE
else
#     echo "script may only be called on files with extensions [$EXTENSIONS], invalid extension \"$extension\" on $FILE"
    exit 1
fi

