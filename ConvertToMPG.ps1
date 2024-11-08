# This script is written to quickly rename and run mass files through a .bat conversion file with ffmpeg.
# a slightly lighter weight version will only rename the files by setting $bRenameOnly to true.
# 
# 
# Before running scripts to enable PowerShell scripts type the following in PowerShell as administrator:
#
# Set-ExecutionPolicy RemoteSigned
#
# Here is an example of !ConvertToMPG.bat in case you need it, see help for ffmpeg for explincation of these options:
#
# ffmpeg -i input.mkv -profile:v baseline -level 3.0 -vf format=yuv420p -preset slow -crf 22 -ac 2 -map_metadata -1 -map 0:0 -map 0:1 output.mp4


# frequently changed options:
#
# trim the prefix for the file names. Often the series name. This will also attempt to identify season and episode numbers
[bool]$bTrimPrefix = $true
# trim the suffix if there are repeate characters after the episode name. i.e. encoding information, etc.
[bool]$bTrimSuffix = $true
# use single digit season numbers. i.e. S01E12 would become S1E12.
[bool]$bSingleDigitSeasonNumbers = $true

# only rename files if they don't need to be re-encoded, but the name need to be cleaned up.
[bool]$bRenameOnly = $false
# if running unattended this option will shut down the PC upon script completion.
[bool]$bShutdownPC = $false

# less used options:

# rename the extention in this function. Only files of this type will be processed.
function GetAllFilesOfType()
{
    # edit the file extention here for other file types:
    return Get-ChildItem -Path . -Filter "*.mkv"
}

# will give useful debug information if the script isn't acting as expected. i.e. return values from functions.
[bool]$bDebugging = $false
# more debugging details, often data about every loop iteration.
[bool]$bVerboseDebugging = $false
# the format the files will be converted into. Note: this must match the output extention in !ConvertToMPG.bat to work properly
[string]$TargetFileExtention = ".mp4"
# the full path of the ffmpeg executable.
[string]$ffmpegPath = "D:\Working\ffmpeg\bin"


# functions
function GetFileExtention([string]$file_name)
{
    [array]$test_array = $file_name.Split(".")
    [int]$index = $test_array.Length - 1
    [string]$return_string = "." + $test_array[$index]

    if ($bDebugging){ Write-Host "GetFileExtention returning:" $return_string }
    return $return_string
}

function GetRepeateString([array]$file_names, [bool]$bFromEnd)
{
    [int]$array_length = $file_names.Length
    if ($array_length -lt 2) { return "" }

    [string]$working_string = $file_names[0]
    [int]$working_length = $working_string.Length

    foreach ($file in $file_names)
    {
        if ($bVerboseDebugging){ Write-Host "FindRepeateChars checking file:" $file }

        [string]$test_string = $file
        [int]$test_length = $test_string.Length
        $working_length = [Math]::Min($working_length, $test_length)

        while($working_length -gt 1)
        {
            if($bFromEnd)
            {
                # .Substring( startIndex, length )
                $working_string = $working_string.Substring(($working_string.Length - $working_length), $working_length)
                $test_string = $test_string.Substring(($test_string.Length - $working_length), $working_length)
            }
            else
            {
                $working_string = $working_string.Substring(0, $working_length)
                $test_string = $test_string.Substring(0, $working_length)
            }

            if ($bVerboseDebugging)
            {
                Write-Host "FindRepeateString while loop Comparing:"
                Write-Host $working_string
                Write-Host $test_string
            }

            if ($working_string -eq $test_string)
            {
                if ($bVerboseDebugging){ Write-Host "Match Found!" }
                break
            }

            $working_length--
        }
    }

    if ($bDebugging){ Write-Host "FindRepeateString returning:" $working_string }
    return $working_string
}

function FindSeasonAndEpisodeNumbers([string]$file_name)
{
    [string]$Season = ""
    [string]$Episode = ""
    [string]$return_string = ""
    [string]$episode_prefix = ""

    [string]$file_name_upper = $file_name.ToUpper()
    [array]$test_array = $file_name_upper.Split("S")

    foreach($test_string in $test_array)
    {
        if ($bVerboseDebugging){ Write-Host "FindSeasonAndEpisodeNumbers foreach loop:" $test_string }
        #Check the string length, the S is gone from the Split() above, so the shortest valid length would be 4, i.e. 1E03 or 12E01
        if (($test_string.Length -ige 4) -and (IsCharNumber $test_string[0]))
        {
            #check for single digit season number
            if (($test_string[1] -eq "E") -and (IsCharNumber $test_string[2]) -and (IsCharNumber $test_string[3]))
            {
                $Season = $test_string.Substring(0, 1)
                $Episode = $test_string.Substring(2, 2)
            }
            #check for double digit season number
            elseif ($test_string.Length -gt 4 -and (IsCharNumber $test_string[1]) -and ($test_string[2] -eq "E") -and (IsCharNumber $test_string[3]) -and (IsCharNumber $test_string[4]))
            {
                $Season = $test_string.Substring(0, 2)
                $Episode = $test_string.Substring(3, 2)
            }
        }
    }

    if ($Season.Length -gt 0 -and $Episode.Length -gt 0)
    {
        [string]$search_string = "S" + $Season + "E" + $Episode
        [int]$index = $file_name_upper.IndexOf($search_string)
        $return_string = $file_name.Substring(($index + $search_string.Length))

        if($bSingleDigitSeasonNumbers -and $Season.Length -gt 1 -and $Season[0] -eq "0")
        {
            $Season = $Season.Substring(1, 1)
        }

        $episode_prefix = "S" + $Season + "E" + $Episode + " - "
    }

    [array]$return_array = @($episode_prefix, $return_string)
    if ($bDebugging){ Write-Host "FindSeasonAndEpisodeNumbers returning:" $return_array }
    return $return_array
}

function TrimePrefix([string]$file_name, [string]$prefix)
{
    [string]$return_string = $file_name

    [int]$index = $file_name.IndexOf($prefix)
    if ($index -gt -1) # -1 means it wasn't found
    {
        [int]$prefix_length = $prefix.Length
        $return_string = $file_name.Substring(($index + $prefix_length), $file_name.Length - $prefix_length)
    }

    if ($bDebugging){ Write-Host "TrimePrefix returning:" $return_string }
    return $return_string
}

function TrimeSuffix([string]$file_name, [string]$suffix)
{
    [string]$return_string = $file_name
    [int]$index = $file_name.IndexOf($suffix)
    if ($index -gt 0)
    {
        $return_string = $file_name.Substring(0, $index)
    }
    else
    {
        #if there isn't a suffix to trim, just trim the original file extention
        [int]$index = $file_name.IndexOf($OriginalFileExtention)
        $return_string = $file_name.Substring(0, $index)
    }

    if ($bDebugging){ Write-Host "TrimeSuffix returning:" $return_string }
    return $return_string
}

function GetCleanFileName([string]$in_string)
{
    [string]$return_string = $in_string

    #replace periods with spaces
    $return_string = $return_string -replace '\.', ' '

    #replace dashes with spaces
    $return_string = $return_string -replace '-', ' '

    # replace & with and, alternatly could add & to IsCharValid, but could possibly add issues to some OS
    # spaces around and are to solve for D&D or simlar and double spaces are cleaned up below.
    $return_string = $return_string -replace '&', ' and '

    #trim leading and trailing spaces
    $return_string = $return_string.Trim()

    [int]$prevLength = $return_string.Length

    for ($i = 0;
        $i -lt $prevLength;
        )
    {
        if((IsCharValid $return_string[$i]))
        {
            $i++
        }
        else
        {
            if ($bVerboseDebugging){ Write-Host "found rejected char:" $return_string[$i] " at index" $i }
            # note: doing this in-line wasn't working, calculating int and strings worked.
            # .Substring( startIndex, length )
            [string]$before = $return_string.Substring(0, $i)
            [int]$start_index = $i + 1
            [int]$length = $prevLength - $start_index
            [string]$after = $return_string.Substring($start_index, $length )
            $return_string = $before + $after
            $prevLength = $return_string.Length
        }
    }

    #remove double spaces, note the +1 is to be sure it runs at least once.
    [int]$prevLength = $return_string.Length + 1
    while($prevLength -ne $return_string.Length)
    {
        $prevLength = $return_string.Length
        $return_string = $return_string -replace '  ', ' '
    }

    if ($bDebugging){ Write-Host "GetCleanFileName returning:" $return_string }
    return $return_string
}

function IsCharValid([char]$char)
{
    if ((IsCharLetter($char)) -or (IsCharNumber($char)))
    {
        return $true
    }
    elseif(($char -eq " ") -or ($char -eq "-") -or ($char -eq ".")) #other valid chars
    {
        return $true
    }

    return $false
}

function IsCharNumber([char]$char)
{
    return [char]::IsNumber($char)
}

function IsCharLetter([char]$char)
{
    return [char]::IsLetter($char)
}


# variables that need to persist
[string]$RepeatePrefix = ""
[string]$RepeateSuffix = ""
[string]$OriginalFileExtention = ""


# program starts here
[array]$FilesArray = GetAllFilesOfType
if ($FilesArray.Length -lt 1)
{
    Write-Host "NO FILES WERE FOUND!"
    break;
}

$OriginalFileExtention = GetFileExtention $FilesArray[0]

if ($bTrimPrefix)
{
    $RepeatePrefix = GetRepeateString $FilesArray $false
}

if ($bTrimSuffix)
{
    $RepeateSuffix = GetRepeateString $FilesArray $true
}

foreach($FileName in $FilesArray)
{
    Write-Host "Looping on file name: " $FileName
    # reset file specific values
    [string]$EpisodePrefix = ""
    [string]$NewFileName = $FileName

    if ($bTrimSuffix)
    {
        $NewFileName = TrimeSuffix $NewFileName $RepeateSuffix
    }

    if ($bTrimPrefix)
    {
        [array]$EpisodeArray = FindSeasonAndEpisodeNumbers $NewFileName
        $EpisodePrefix = $EpisodeArray[0]
        if ($EpisodePrefix.Length -gt 0)
        {
            $NewFileName = $EpisodeArray[1]
        }
        else
        {
            $NewFileName = TrimePrefix $NewFileName $RepeatePrefix
        }
    }

    $NewFileName = GetCleanFileName $NewFileName

    # put the name back togather $EpisodePrefix + $NewFileName + extention
    if ($bRenameOnly)
    {
        $NewFileName = $EpisodePrefix + $NewFileName + $OriginalFileExtention
        Rename-Item -literalpath $FileName $NewFileName
    }
    else
    {
        [string]$WorkingLocation = Get-Location
        $NewFileName = $EpisodePrefix + $NewFileName + $TargetFileExtention
        [string]$CopyTo = $ffmpegPath + "\input" + $OriginalFileExtention
        Copy-Item -literalpath $FileName $CopyTo
        Set-Location $ffmpegPath

        cmd.exe /c '!ConvertToMPG.bat'
        $CopyTo = $WorkingLocation + "\" + $NewFileName
        Copy-Item -literalpath "output.mp4" $CopyTo
        [string]$inputFileName = "input" + $OriginalFileExtention
        Remove-Item $inputFileName
        Remove-Item "output.mp4"

        Set-Location $WorkingLocation
    }

    Write-Host "Renamed: " $FileName 
    Write-Host "     to: " $NewFileName
}

if ($bShutdownPC)
{
   # shutdown /s
    Stop-Computer -ComputerName localhost
}
