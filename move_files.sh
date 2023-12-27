#!/bin/bash

# Create the 'lib' folder if it doesn't exist
mkdir -p lib

# Move files starting with underscore into the 'lib' folder
mv _* lib/

# Move folders named 'tcl' and 'tk' into the 'lib' folder
mv tcl lib/
mv tk lib/

# Move files ending with .pyd into the 'lib' folder
mv *.pyd lib/

# Move all files ending in .dll unless they start with 'python3' into the 'lib' folder
shopt -s extglob
mv !(python3*).dll lib/

# Delete the script file itself
rm "$0"