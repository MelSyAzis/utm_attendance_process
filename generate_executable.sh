#!/bin/bash

SCRIPT_NAME=process_attendance.py
EXECUTABLE_NAME="${SCRIPT_NAME%.*}"

pyinstaller --onefile $SCRIPT_NAME
mkdir -p bin
mv dist/$EXECUTABLE_NAME bin/$EXECUTABLE_NAME