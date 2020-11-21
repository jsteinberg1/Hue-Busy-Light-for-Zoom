# Hue-Busy-Light-for-Zoom
Turn on a hue light based on your Zoom status

## Python Script
1. Install python 3.8
2. Install dependencies from requirements.txt ( e.g. pip install -r requirements.txt )
3. Run python script

Note: To terminate the script when run as a background process, kill the python process using Windows task manager.

# Windows EXE files

Windows EXE files are provided in the dist folder.

## Standlone EXE

1. Run the 'Hue Busy Light For Zoom.exe' file.

## Run as a Windows service

1. Run the Standalone exe file first to initialize the connection to the Hue bridge and validate functionality.
2. Once the standalone exe is working, you can install the Busy Light as a service using 'Hue_Busy_Light_for_Zoom_service.exe install'
3. The Busy Light service currently needs to 'Log On As' the same Windows user account that performed step 1 above. 

## Configuration Settings

The Hue settings are stored in HKCU\Software\Hue Busy Light for Zoom
