#!/bin/bash 

echo "Deleting old ciCHECK files"
DIREC=/Users/Shared/cicheckFiles
FILE1=/Applications/ciCHECK.app

killall ciCHECK

if [ -d $DIREC ]
then
	sudo rm -R /Users/Shared/cicheckFiles 
fi

if [ -d $FILE1 ]
then
	sudo rm -rf /Applications/ciCHECK.app
fi

