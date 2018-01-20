<#
.SYNOPSIS
MediaDescGen prepares a video file with a Mediainfo description and screenshots
.DESCRIPTION
This script first ensures it's dependencies (Mediainfo, Avinaptic, and ffmpeg) are downloaded and available.
Next, it takes all video files in a given directory and moves each file into its own directory. 
It then generates a description for the file based on certain Mediainfo fields, and finally takes 2 screenshots of the video.
.PARAMETER directory
The directory to search for video files in. Defaults to the folder the script is running from.
.EXAMPLE
./MediaDescGen.ps1 -directory "C:\Users\User\Downloads"
#>

param (
    [string]$directory = "$PSScriptRoot"
)

# Define utility paths
$mediainfo_path = "$PSScriptRoot\MediaDescGen_Utilities\mediainfo\MediaInfo.exe"
$avinaptic_path = "$PSScriptRoot\MediaDescGen_Utilities\avinaptic\avinaptic2-20111218\avinaptic2-cli.exe"
$ffmpeg_path = "$PSScriptRoot\MediaDescGen_Utilities\ffmpeg\ffmpeg-20180102-57d0c24-win64-static\bin\ffmpeg.exe"

# Download utilities, if needed
if (!(Test-Path $PSScriptRoot/MediaDescGen_Utilities)) {
    Write-Host "Utilities folder not found, creating..."
    mkdir $PSScriptRoot/MediaDescGen_Utilities | Out-Null
}
if (!(Test-Path $mediainfo_path)) {
    Write-Host "Mediainfo not found, downloading..."
    Invoke-WebRequest https://mediaarea.net/download/binary/mediainfo/17.12/MediaInfo_CLI_17.12_Windows_x64.zip -OutFile $PSScriptRoot/MediaDescGen_Utilities/mediainfo.zip
    Expand-Archive $PSScriptRoot/MediaDescGen_Utilities/mediainfo.zip -DestinationPath $PSScriptRoot/MediaDescGen_Utilities/mediainfo
    Remove-Item $PSScriptRoot/MediaDescGen_Utilities/mediainfo.zip
    Write-Host "Mediainfo downloaded."
}
if (!(Test-Path $avinaptic_path)) {
    Write-Host "Avinaptic not found, downloading..."
    Invoke-WebRequest http://fsinapsi.altervista.org/code/avinaptic/avinaptic2-win32-20111218.zip -OutFile $PSScriptRoot/MediaDescGen_Utilities/avinaptic.zip
    Expand-Archive $PSScriptRoot/MediaDescGen_Utilities/avinaptic.zip -DestinationPath $PSScriptRoot/MediaDescGen_Utilities/avinaptic
    Remove-Item $PSScriptRoot/MediaDescGen_Utilities/avinaptic.zip
    Write-Host "Avinaptic downloaded."
}
if (!(Test-Path $ffmpeg_path)) {
    Write-Host "ffmpeg not found, downloading..."
    Invoke-WebRequest https://ffmpeg.zeranoe.com/builds/win64/static/ffmpeg-20180102-57d0c24-win64-static.zip -OutFile $PSScriptRoot/MediaDescGen_Utilities/ffmpeg.zip
    Expand-Archive $PSScriptRoot/MediaDescGen_Utilities/ffmpeg.zip -DestinationPath $PSScriptRoot/MediaDescGen_Utilities/ffmpeg
    Remove-Item $PSScriptRoot/MediaDescGen_Utilities/ffmpeg.zip
    Write-Host "ffmpeg downloaded."
}

function buildDesc {
    param (
        $file_path
    )
    # Get path of parent directory
    $directory = $file_path.Directory.FullName
    # Get file name without extension
    $file = $file_path.Name
    $name = [io.path]::GetFileNameWithoutExtension("$file")
    # Move file into its own directory
    Write-Host "Moving $file into its own directory..."
    mkdir $directory/$name | out-null
    mv $directory/$file $directory/$name/$file
    # Check for an .nfo file with the same name and, if it exists, move it too
    if (Test-Path $directory/$name.nfo) {
        mv $directory/$name.nfo $directory/$name/
    }

    Write-Host "Creating description..."
    ###############
    ### General ###
    ###############
    Add-Content -Path $directory/$name/$name-description.txt -Value "[b]General[/b]"
    # Get the name of the file
    Add-Content -Path $directory/$name/$name-description.txt -Value "Name: $name"
    # Get the video container
    [IO.File]::AppendAllText("$directory/$name/$name-description.txt","Format: ")
    cmd /c $mediainfo_path --output=General`;%Format% $directory/$name/$file | Add-Content $directory/$name/$name-description.txt
    # Get the duration in ms and convert to HH:MM:SS
    [IO.File]::AppendAllText("$directory/$name/$name-description.txt","Duration: ")
    $duration =  cmd /c $mediainfo_path --output=General`;%Duration% $directory/$name/$file 
    [timespan]::FromMilliseconds($duration).ToString("hh\:mm\:ss") | Add-Content $directory/$name/$name-description.txt

    #############
    ### VIDEO ###
    #############
    Add-Content -Path $directory/$name/$name-description.txt -Value "`r`n[b]Video[/b]"
    # Get the video codec
    [IO.File]::AppendAllText("$directory/$name/$name-description.txt","Codec: ")
    cmd /c $mediainfo_path --output=Video`;%Format% $directory/$name/$file | Add-Content $directory/$name/$name-description.txt
   
    # If the bitrate is variable, get the max bitrate
    if ((cmd /c $mediainfo_path --output=Video`;%BitRate_Mode/String% $directory/$name/$file) -like "Variable*") {
        [IO.File]::AppendAllText("$directory/$name/$name-description.txt","Bitrate mode: Variable`r`n")
        [IO.File]::AppendAllText("$directory/$name/$name-description.txt","Maximum Bitrate (mbps): ")
        $videobitrate = cmd /c $mediainfo_path --output=Video`;%BitRate_Maximum% $directory/$name/$file
        ($videobitrate/1mb).Tostring(".000") | Add-Content $directory/$name/$name-description.txt
    }
    else {
    # Get the bitrate in bps, and convert to Mbps
        [IO.File]::AppendAllText("$directory/$name/$name-description.txt","Bitrate mode: Constant`r`n")
        [IO.File]::AppendAllText("$directory/$name/$name-description.txt","Bitrate (mbps): ")
        $videobitrate = cmd /c $mediainfo_path --output=Video`;%BitRate% $directory/$name/$file
        ($videobitrate/1mb).Tostring(".000") | Add-Content $directory/$name/$name-description.txt
    }
    # Get the horizontal resolution
    [IO.File]::AppendAllText("$directory/$name/$name-description.txt","Width: ")
    cmd /c $mediainfo_path --output=Video`;%Width% $directory/$name/$file | Add-Content $directory/$name/$name-description.txt
    # Get the vertical resolution
    [IO.File]::AppendAllText("$directory/$name/$name-description.txt","Height: ")
    cmd /c $mediainfo_path --output=Video`;%Height% $directory/$name/$file | Add-Content $directory/$name/$name-description.txt
    # Get the video frame rate
    [IO.File]::AppendAllText("$directory/$name/$name-description.txt","Framerate: ")
    cmd /c $mediainfo_path --output=Video`;%FrameRate% $directory/$name/$file | Add-Content $directory/$name/$name-description.txt
    
    #############
    ### AUDIO ###
    #############
    Add-Content -Path $directory/$name/$name-description.txt -Value "`r`n[b]Audio[/b]"
    [IO.File]::AppendAllText("$directory/$name/$name-description.txt","Codec: ")
    cmd /c $mediainfo_path --output=Audio`;%Format% $directory/$name/$file | Add-Content $directory/$name/$name-description.txt
    # Get bitrate from avinaptic if AAC audio
    if ((cmd /c $mediainfo_path --output=Audio`;%Format% $directory/$name/$file) -like "*AAC*") {
        # Get avinaptic report
        $avinaptic = cmd /c $avinaptic_path "$directory/$name/$file" 2> $null
        # Transform report into an array
        $avinaptic = $avinaptic -split "`n"
        $avinaptic = $avinaptic -match '[ Audio track ]'
        # Trim down to only the audio information
        $audio_position = [array]::IndexOf($avinaptic,"[ Audio track ]")
        $avinaptic = $avinaptic[$audio_position..$avinaptic.Length]
        # Get the bitrate from the audio information
        $audiobitrate = $avinaptic -like "Bitrate*"
        # If this returns no result, the bitrate is variable, and the track requires a full analysis with Avinaptic
        if ($audiobitrate.Length -lt 1) {
            echo "The audio track requires full analysis, this may take a minute..."
            # Get avinaptic report
            $avinaptic = cmd /c $avinaptic_path --drf "$directory/$name/$file" 2> $null
            # Transform report into an array
            $avinaptic = $avinaptic -split "`n"
            $avinaptic = $avinaptic -match '[ Audio track ]'
            # Trim down to only the audio information
            $audio_position = [array]::IndexOf($avinaptic,"[ Audio track ]")
            $avinaptic = $avinaptic[$audio_position..$avinaptic.Length]
            # Get the bitrate from the audio information
            $audiobitrate = $avinaptic -like "Bitrate*"
        }
        # Write to description
        Add-Content $directory/$name/$name-description.txt -Value $audiobitrate
    }
    # Else fall back to mediainfo
    else {
        # Get the audio bitrate in bps and convert to kbps
        [IO.File]::AppendAllText("$directory/$name/$name-description.txt","Bitrate (kbps): ")
        $audiobitrate = cmd /c $mediainfo_path --output=Audio`;%BitRate% $directory/$name/$file
        ($audiobitrate/1kb).Tostring(".000") | Add-Content $directory/$name/$name-description.txt
    }
    # Get the number of audio channels
    [IO.File]::AppendAllText("$directory/$name/$name-description.txt","Channels: ")
    cmd /c $mediainfo_path --output=Audio`;%Channels% $directory/$name/$file | Add-Content $directory/$name/$name-description.txt
    # Get the audio language
    [IO.File]::AppendAllText("$directory/$name/$name-description.txt","Language: ")
    cmd /c $mediainfo_path --output=Audio`;%Language/String% $directory/$name/$file | Add-Content $directory/$name/$name-description.txt
    
    #################
    ### SUBTITLES ###
    #################
    Add-Content -Path $directory/$name/$name-description.txt -Value "`r`n[b]Subtitles[/b]"
    # Get subtitle format
    [IO.File]::AppendAllText("$directory/$name/$name-description.txt","Format: ")
    if ((cmd /c $mediainfo_path --output=Text`;%Format% $directory/$name/$file).Length -lt "2") {
        Add-Content $directory/$name/$name-description.txt -Value "None`n"
    }
    else {
        cmd /c $mediainfo_path --output=Text`;%Format% $directory/$name/$file | Add-Content $directory/$name/$name-description.txt
    }
    # Get subtitle language
    [IO.File]::AppendAllText("$directory/$name/$name-description.txt","Language: ")
    if ((cmd /c $mediainfo_path --output=Text`;%Language/String% $directory/$name/$file).Length -lt "2") {
        Add-Content $directory/$name/$name-description.txt -Value "None`n"
    }
    else {
        cmd /c $mediainfo_path --output=Text`;%Language/String% $directory/$name/$file | Add-Content $directory/$name/$name-description.txt
    }

    # Get screenshots
    Write-Host "Description done; taking screenshot #1 at 5 minutes in..."
    cmd /c $ffmpeg_path -ss 00:05:00 -t 1 -i $directory/$name/$file $directory/$name/$name-screenshot1.png 2> $null
    Write-Host "Screenshot 1 done; taking screenshot 2 at 20 minutes in..."
    cmd /c $ffmpeg_path -ss 00:20:00 -t 1 -i $directory/$name/$file $directory/$name/$name-screenshot2.png 2> $null
    Write-Host "Screenshot 2 done."
    
    # Done
    Write-Output "Prepared description and screenshots for $name."
}

# Get a list of all video files in the directory
foreach ($item in (Get-ChildItem $directory/* -include "*.m4v","*.mkv","*.mp4")) {
    buildDesc $item
}