#!/bin/bash

SCRIPT_NAME=main.py
EXECUTABLE_NAME=process_attendance

pyinstaller --onefile --name $EXECUTABLE_NAME $SCRIPT_NAME
mkdir -p bin
mv dist/$EXECUTABLE_NAME bin/$EXECUTABLE_NAME