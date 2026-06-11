#!/bin/sh
INSTALL_MANIFEST=install_manifest.txt

if [ "$1" = "doit" ]; then
	if [ "`id -u`" -ne 0 ]; then
		echo "`basename $0`: doit action must be run as root / with sudo"
		exit
	fi

	if [ -e "$INSTALL_MANIFEST" ]; then
		echo "Uninstalling OpenXLSX per $INSTALL_MANIFEST..."
		FILES=`cat $INSTALL_MANIFEST`
		for FILE in $FILES; do
			if [ -e $FILE ] || [ -L $FILE ]; then
				if [ -e $FILE ]; then
					echo "-- Uninstalling: $FILE"
				else
					echo "-- Removing symbolic link: $FILE"
				fi
				rm $FILE
			else
				echo "-- Already uninstalled: $FILE"
			fi
		done

		DIRECTORIES=`cat $INSTALL_MANIFEST | xargs dirname | sort --reverse | uniq`
		for DIRECTORY in $DIRECTORIES; do
			if [ -e "$DIRECTORY" ]; then
				DIRLIST=`ls -A "$DIRECTORY"`
				if [ -z "$DIRLIST" ]; then # test if directory is empty
					echo "-- Removing directory: $DIRECTORY"
					rmdir -p --ignore-fail-on-non-empty "$DIRECTORY" # NOTE: must still ignore fail otherwise rmdir -p will exit with error on the first parent dir that is not empty
				else
					echo "-- Directory not empty: $DIRECTORY"
				fi
			else
				echo "-- Already removed: $DIRECTORY"
			fi
		done
	else
		echo "Could not find $INSTALL_MANIFEST - can not uninstall OpenXLSX"
	fi
else
	echo "--- Uninstall using the following commands: ---"
	echo "cat $INSTALL_MANIFEST | sort | uniq | sudo xargs rm"
	echo "cat $INSTALL_MANIFEST | xargs dirname | sort | uniq | sudo xargs rmdir -p"
	# echo "cat $INSTALL_MANIFEST | xargs dirname | sort | uniq | sudo xargs rmdir -p --ignore-fail-on-non-empty"
fi
