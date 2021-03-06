<#	
	.NOTES
	===========================================================================
	 Created with: 	Visual Studio Code 1.51.1
	 Created on:   	11/27/2020 1:09 PM
	 Created by:   	Jason Witters
	 Organization: 	Witters Inc.
	 Filename:     	Compare-MovieFiles.ps1
	===========================================================================
	.DESCRIPTION
		Compares media files between directories and copies over new versions of existing movies.
#>

[CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $false,
            Position = 0,
            HelpMessage = "Directory path of new movie downloads")]
        [string]$NewMovieDirectory = "Z:\Film\_New\Bluray",

        [Parameter(
            Mandatory = $false,
            Position = 1,
            HelpMessage = "Directory path of final movie folder")]
        [string]$CurrentMovieDirectory = "Z:\Film\Movies",

        [Parameter(
            Mandatory = $false,
            Position = 2)]
        [switch]$Test = $false
    )

#ScriptVersion = "1.0.2.0"

$old = Get-ChildItem -Path $CurrentMovieDirectory
$new = Get-ChildItem -Path $NewMovieDirectory
$Results = Compare-Object -ReferenceObject $new -DifferenceObject $old -IncludeEqual -ExcludeDifferent -Property Name
if ($Results)
{
    Write-Host "Test Mode: $Test" -ForegroundColor Blue
    Write-Host "Number of matches found: $($Results.count)" -ForegroundColor Blue
    Write-Host "---------" -ForegroundColor Blue
    "`n"
    $Results.Name
    "`n"
    Write-Host "Copy new versions to Movie directory ($CurrentMovieDirectory)?" -ForegroundColor Magenta
    $Answer = Read-Host "Y/N?"
    if ($Answer -match "[Yy]")
    {
        foreach ($Result in $Results.Name)
        {
            Write-Host "Processing: $Result" -ForegroundColor Blue
            
            # New and Old movie directories
            $TempMovieDirectory = Join-Path -Path $NewMovieDirectory -ChildPath $Result
            $FinalMoviePath = Join-Path -Path $CurrentMovieDirectory -ChildPath $Result

            # Get child items for each movie directory for comparison
            $NewMovieFiles = Get-ChildItem $TempMovieDirectory
            $ExistingMovieFiles = Get-ChildItem $FinalMoviePath

            $MovieComparison = Compare-Object -ReferenceObject $NewMovieFiles -DifferenceObject $ExistingMovieFiles -Property Name -IncludeEqual
            $MKVMoviePath = Join-Path -Path $FinalMoviePath -ChildPath "$Result.mkv"
            $MP4MoviePath = Join-Path -Path $FinalMoviePath -ChildPath "$Result.mp4"
            $AVIMoviePath = Join-Path -Path $FinalMoviePath -ChildPath "$Result.avi"
            $SRTMoviePath = Join-Path -Path $FinalMoviePath -ChildPath "$Result.srt"
            $NFOMoviePath = Join-Path -Path $FinalMoviePath -ChildPath "$Result.nfo"
            $BIFMoviePath = Join-Path -Path $FinalMoviePath -ChildPath "$Result-320-10.bif"

            foreach ($MovieItem in $MovieComparison)
            {
                $MovieItemName = $MovieItem.Name
                $MovieItemTempPath = Join-Path -Path $TempMovieDirectory -ChildPath $MovieItemName
                Write-Host "Processing file: $MovieItemName" -ForegroundColor Blue
                try
                {
                    # If file exists in both folders, copy new version over
                    if ($MovieItem.SideIndicator -like "==")
                    {
                        try
                        {
                            Write-Host "File match, copying file over..." -ForegroundColor Blue
                            Move-Item -Path $MovieItemTempPath -Destination $FinalMoviePath -Force -ErrorAction Stop -WhatIf:$Test
                            Write-Host "SUCCESS! Moved file ($MovieItemTempPath) to ($FinalMoviePath)" -ForegroundColor Green
                        }
                        catch
                        {
                            Write-Host "Error encountered copying file ($MovieItemName)" -ForegroundColor Red
                            exit
                        }
                    }
                    # If file exists only in new folder
                    elseif ($MovieItem.SideIndicator -like "<=")
                    {
                        # 4K version
                        if ($MovieItemName -like "$Result - 4K*")
                        {
                            Write-Host "New 4K version detected, copy over and rename existing files?" -ForegroundColor Magenta
                            $Answer2 = Read-Host "Y/N?"
                            if ($Answer2 -match "[Yy]")
                            {
                                try
                                {
                                    Write-Host "Copying over 4K file..." -ForegroundColor Blue
                                    Move-Item -Path $MovieItemTempPath -Destination $FinalMoviePath -Force -ErrorAction Stop -WhatIf:$Test
                                    Write-Host "SUCCESS! Moved file ($MovieItemTempPath) to ($FinalMoviePath)" -ForegroundColor Green
                                    if (Test-Path $MP4MoviePath)
                                    {
                                        Write-Host "Renaming file: $MP4MoviePath" -ForegroundColor Blue
                                        $Rename = Rename-Item -Path $MP4MoviePath -NewName "$Result - HD.mp4" -Force -PassThru -ErrorAction Stop -WhatIf:$Test
                                        Write-Host "SUCCESS! Renamed file: $MP4MoviePath" -ForegroundColor Green
                                    }
                                    elseif (Test-Path $MKVMoviePath)
                                    {
                                        Write-Host "Renaming file: $MKVMoviePath" -ForegroundColor Blue
                                        $Rename = Rename-Item -Path $MKVMoviePath -NewName "$Result - HD.mkv" -Force -PassThru -ErrorAction Stop -WhatIf:$Test
                                        Write-Host "SUCCESS! Renamed file: $MKVMoviePath" -ForegroundColor Green
                                    }
                                    # if 4k copy and HD rename success, handle other files
                                    if ($Rename)
                                    {
                                        # Rename subtitle file
                                        if (Test-Path $SRTMoviePath)
                                        {
                                            Write-Host "Renaming file: $SRTMoviePath" -ForegroundColor Blue
                                            Rename-Item -Path $SRTMoviePath -NewName "$Result - HD.srt" -Force -ErrorAction Stop -WhatIf:$Test
                                            Write-Host "SUCCESS! Renamed file: $SRTMoviePath" -ForegroundColor Green
                                        }
                                        # Delete .NFO file
                                        if (Test-Path $NFOMoviePath)
                                        {
                                            Write-Host "Removing file: $NFOMoviePath" -ForegroundColor Blue
                                            Remove-Item -Path $NFOMoviePath -Force -Confirm:$false -WhatIf:$Test
                                            Write-Host "SUCCESS! Removed file: $NFOMoviePath" -ForegroundColor Green
                                        }
                                        # Delete .BIF File
                                        if (Test-Path $BIFMoviePath)
                                        {
                                            Write-Host "Removing file: $BIFMoviePath" -ForegroundColor Blue
                                            Remove-Item -Path $BIFMoviePath -Force -Confirm:$false -WhatIf:$Test
                                            Write-Host "SUCCESS! Removed file: $BIFMoviePath" -ForegroundColor Green
                                        }
                                    }
                                }
                                catch
                                {
                                    Write-Host "Error encountered" -ForegroundColor Red
                                }
                            }
                        }
                        # Non 4k version
                        else
                        {
                            # Skip interaction if relacing .MKV file with .MP4 file
                            if ($MovieItemName -like "$Result.mp4")
                            {
                                if (Test-Path $MKVMoviePath)
                                {
                                    $RemoveThis = $MKVMoviePath
                                }
                                elseif (Test-Path $AVIMoviePath)
                                {
                                    $RemoveThis = $AVIMoviePath
                                }
                                try
                                {
                                    Write-Host "Copying new .MP4 file over existing .MKV/.AVI file" -ForegroundColor Blue
                                    Remove-Item -Path $RemoveThis -Force -ErrorAction Stop -WhatIf:$Test
                                    Write-Host "SUCCESS! Removed file: $RemoveThis" -ForegroundColor Green
                                    $Answer2 = "y"
                                }
                                catch
                                {
                                    Write-Host "Error encountered copying file ($MovieItemName)" -ForegroundColor Red
                                    exit
                                }
                            }
                            # Skip interaction if subtitle file
                            elseif ($MovieItemName -like "$Result.srt")
                            {
                                $Answer2 = "y"
                            }
                            else
                            {
                                Write-Host "New file found: ($MovieItemName)" -ForegroundColor Blue
                                Write-Host "Existing movie files:" -ForegroundColor Blue
                                foreach ($File in $ExistingMovieFiles)
                                {
                                    Write-Host $File.FullName -ForegroundColor Blue
                                }
                                Write-Host "Copy over new item ($MovieItemName)?" -ForegroundColor Magenta
                                $Answer2 = Read-Host "Y/N?"
                            }

                            if ($Answer2 -match "[Yy]")
                            {
                                try
                                {
                                    Write-Host "Permission granted, copying file over" -ForegroundColor Green
                                    Move-Item -Path $MovieItemTempPath -Destination $FinalMoviePath -Force -ErrorAction Stop -WhatIf:$Test
                                    Write-Host "SUCCESS! Moved file: $MovieItemTempPath" -ForegroundColor Green
                                }
                                catch
                                {
                                    Write-Host "Error encountered copying file ($MovieItemName)" -ForegroundColor Red
                                    exit
                                }
                            }
                            else
                            {
                                Write-Host "Permission denied or incorrect response ($Answer2)"
                            }
                        }
                    }
                }
                catch
                {
                    Write-Warning "Something went wrong!"
                    exit
                }
            }
            
            # Remove temp movie folder if empty
            $NewMovieFiles = Get-ChildItem $TempMovieDirectory
            $FileCount = ($NewMovieFiles | Measure-Object).Count
            if ($FileCount -eq 0)
            {
                try
                {
                    Write-Host "Removing temp folder: ($TempMovieDirectory)" -ForegroundColor Blue
                    Remove-Item -Path $TempMovieDirectory -Force -Recurse -ErrorAction Stop -WhatIf:$Test
                    Write-Host "SUCCESS! Removed directory ($TempMovieDirectory" -ForegroundColor Green
                }
                catch
                {
                    Write-Warning "Something went wrong!"
                    exit
                }
            }
        }
    }
}
else
{
    Write-Host "No matches found"
}