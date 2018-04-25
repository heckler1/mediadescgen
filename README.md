# MediaDescGen

MediaDescGen prepares the necessary files to upload a Movie/TV Episode to the web. Given a directory, it moves each video file in the directory into its own directory, and then generates a description and screenshots of each file. If no directory is supplied, the current directory of the script is used.

Usage: `.\MediaDescGen.ps1 -directory "C:\Users\User\Downloads\"`

On the first run, MediaDescGen downloads the required utilities to a folder named "MediaDescGen_Utilities", created in the same directory that the script is run from.

This script uses MediaInfo, and if necessary Avinaptic, to gather the required information about the file and format it into a description, with BB code section headers. It also uses ffmpeg to take screenshots of the file at 5mins and 20mins in.



