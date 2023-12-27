#!/bin/bash

# Create the 'lib' folder if it doesn't exist
mkdir -p lib

# Move files starting with underscore into the 'lib' folder
mv _* lib/

# Move folders named 'tcl' and 'tk' into the 'lib' folder
mv tcl lib/
mv tcl8 lib/
mv tk lib/
mv certifi lib/
mv google_api_core-1.34.0.dist-info lib/
mv google_api_python_client-1.8.0.dist-info lib/
mv httplib2 lib/

# Move files ending with .pyd into the 'lib' folder
mv *.pyd lib/

# Move all files ending in .py into the 'lib' folder
mv *.py lib/

# Move all files ending in .dll unless they start with 'python3' into the 'lib' folder
shopt -s extglob
mv !(python3*).dll lib/

# Delete the script file itself
rm "$0"