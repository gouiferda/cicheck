#!/bin/bash 

echo "Making the ciCHECK Files writable..."
DIREC=/Users/Shared/cicheckFiles
FILE1=/Applications/ciCHECK.app

# if the file doesn't exist, try to create folder
if [ -d $DIREC ]
then
 	find $DIREC  -type d -exec chmod 777 {} \;
	find /Users/Shared/cicheckFiles -type f -exec chmod 777 {} \;
fi

if [ -d $FILE1 ]
then
	sudo chmod 755 /Applications/ciCHECK.app
	sudo chmod +x /Applications/ciCHECK.app/Contents/MacOS/ciCHECK
fi
 