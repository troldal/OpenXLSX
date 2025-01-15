#!/bin/sh

if [ -z $1 ]; then
	echo "need to provide a command line parameter"
	exit 1
fi

if [ ! -e $1 ]; then
	echo "provided argument \"$1\" does not exist"
	exit 1
fi

# extension=`echo $1 | sed -r 's/[^.]*(\.([^.]*))*$/\2/g'`
#if [ "$extension" != "xml" ]; then
#	echo "script may only be called on xml files, invalid extension \"$extension\""
#	exit 1
#fi

if [ ! -e $1 ]; then
	echo "provided argument \"$1\" does not exist"
	exit 1
fi

# Old instruction for echo:
# echo "XMLLINT_INDENT=\$'\t' xmllint --format \"$1\" > \"$1.formatted\"; mv \"$1.formatted\" \"$1\""

# New instructions: echo and execute
echo "export XMLLINT_INDENT='	'; xmllint --format \"$1\" > \"$1.formatted\"; mv \"$1.formatted\" \"$1\""

export XMLLINT_INDENT='	';
xmllint --format "$1" > "$1.formatted";

# Safeguard: do not overwrite original file unless xmllint succeeded
if [ "$?" = "0" ]; then
  mv "$1.formatted" "$1"
else
  rm "$1.formatted"
fi
