#!/bin/sh

if [ -z $1 ]; then
	echo "need to provide a command line parameter"
	exit 1
fi

if [ ! -e $1 ]; then
	echo "provided argument \"$1\" does not exist"
	exit 1
fi

SCRIPT_PATH=`dirname "$0"`
# echo "SCRIPT_PATH is $SCRIPT_PATH"

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
