#
# Usage (scan folder is recursive)--
# powershell -f "Gamecube ISO Organizer.ps1" -scan "e:\PATH\TO\ISOS_AND_GCMS\TO_SCAN" -dest "e:\PATH\TO\DESTINATION"
#
# -scan can be filtered to a single game or wildcard match
#
# v1.1: converted to powershell due to file names and special characters, also reverted back to discex v0.8b due to disc 2 detection issues
#

#TODO game.iso and disk2.iso could potentially get reversed if for some reason (disk disconnect) happens and causes disk2 to be named game.iso
Param(
	[String]$scan = "",
	[String]$dest = "",
	[String]$discex = "$PSScriptRoot\DiscEx",
	[String]$gcit = "$PSScriptRoot\GC_ISO_Tool",
	[String]$throttle = 60, 
	[Switch]$md = $False
)

Function EscapePath($path) {
	return $path.Replace('[','`[').Replace(']','`]')
}

function WriteError($err) {
	filter timestamp {"$(Get-Date -Format G): $_"}
	write-host $err
	$err | timestamp | Add-Content "$PSScriptRoot\_error.txt"
}

Function Remove-InvalidFileNameChars {
  param(
    [Parameter(Mandatory=$true,
      Position=0,
      ValueFromPipeline=$true,
      ValueFromPipelineByPropertyName=$true)]
    [String]$Name
  )

  $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
  $re = "[{0}]" -f [RegEx]::Escape($invalidChars)
  return ($Name -replace $re)
}


function DoGame ($iso) {
	Set-Location $discex
	
	write-host ("Reading " + ('"' + $iso + '"') + " for GameID and Title...")
	$arg = @("-v", ('"' + $iso + '"'))
	$output = & "$discex\DiscEx.exe" $arg | Out-String
	
	$gameid = ""
	$title = ""
	$foldername = ""
	$filename = [System.IO.Path]::GetFileNameWithoutExtension("$iso")
	#todo make a better case insenstive replace
	$filename = $filename -replace '\((D|d)isc \d*\)','(multi-disc)'

	
	$lines = $output.split("`r`n")
	foreach ($line in $lines) {
		if ($line -match ":") {
			$header = $line.Substring(0, $line.IndexOf(":"))
			$value = $line.Substring($line.IndexOf(":")+1)
			if ($header -eq "GameID") {
				$gameid = $value.Trim()
			} elseif ($header -eq "Title") {
				$title = $value.Trim()
			}
		}
		
		if ($gameid -and $title) { break }
		#if (($line -notmatch "Copying") -and $line.Trim() -ne "") { write-host $line }
	}
	if (-not $gameid -or -not $title) {
		WriteError -err "Error reading game info from '$iso'"
		return 1
	}

	$preserve_filename = $True
	if($preserve_filename){
		$foldername = Remove-InvalidFileNameChars -name "$filename [$gameid]"
	}else{
		$foldername = Remove-InvalidFileNameChars -name "$title [$gameid]"
	}
	#write-host $foldername
	
	if (test-path -literalpath "$discex\$gameid") {
		remove-item -literalpath "$discex\$gameid" -recurse -force
	}
	
	$out_folder = ($dest + "\$foldername")
	$out_iso = "game.iso"
	
	#todo support more than 2 disc multidiscs
	$disc1_folder = MultiDisc_CheckFolder -folder $out_folder -check_only $True
	if ($disc1_folder) {
		$out_folder = $disc1_folder
		$out_iso = "disc2.iso"
	}
	
	if (test-path -literalpath "$out_folder\$out_iso") {
		Write-Host "Already Exists: '$out_folder\$out_iso'"
		return 0
	}
	
	if (-not (test-path -literalpath "$out_folder")) {
		$ret = New-Item "$out_folder" -type directory
		if (-not (test-path -literalpath "$out_folder")) {
			Write-Host "Error creating '$out_folder'"
			return 0
		}
	}
	
	$move_only = $True
	if ($move_only) {
		write-host "Moving to '$out_folder\$out_iso'..."
		Move-Item -Path "$iso" -Destination "$out_folder\$out_iso"
	}else{
		write-host "Creating '$out_folder\$out_iso'..."
		Set-Location $gcit
		$arg = @("""$iso""", "-AQ", "-F", "FullISO", "-D", """$out_folder\$out_iso""")
	  #& "$gcit\gcit.exe" $arg
	
	  $proc = start-process -PassThru "$gcit\gcit.exe" ($arg -join " ")
	  $proc.WaitForExit()
	}
	
	if (-not (test-path -literalpath "$out_folder\$out_iso")) {
		return 1
	}
	
	return 0
}

function MultiDisc_Check ($folder) {
	$check_folders = Get-ChildItem -Path $folder | ?{ $_.PSIsContainer } | Sort
	foreach ($folder in $check_folders) {
		$ret = MultiDisc_CheckFolder -folder $folder.FullName -check_only $False
	}
}

function MultiDisc_CheckFolder ($folder, $check_only = $True) {
	$folder_name = $folder.Substring($folder.LastIndexOf("\")+1).trim()
	if ($folder_name.length -lt 10) { return 0 }
	
	$gameid = $folder_name.substring($folder_name.length-9).trim()
	if ($gameid.substring(0,1) -ne "[" -or $gameid.substring($gameid.length-1,1) -ne "]") {
		$gameid = $folder_name.substring($folder_name.length-8).trim()
		if ($gameid.substring(0,1) -ne "[" -or $gameid.substring($gameid.length-1,1) -ne "]") {
			WriteError ("Notice: Error determining gameid for '" + $folder + "'")
			return 0
		}
	}
	$gameid = $gameid.substring(1,$gameid.length-2)
	if ($gameid.length -lt 7) { return 0 }
	
	$title = $folder_name.substring(0,$folder_name.length-$gameid.length-2).trim()
	
	#$gametitle = "$title [$gameid]"
	$disc_one_id = ("[" + $gameid.Substring(0,6) + "]")
	
	$check_files = Get-ChildItem -Path $dest | Where-Object { $_.Name -match ([Regex]::Escape($disc_one_id) + "$") }
	if ($check_files.count -gt 1) {
		WriteError -err "More than one folder match on multi-disc check: '$disc_one_id'"
		return 0
	} elseif ($check_files.count -eq 0) {
		write-host "Disc 2 of Multi-Disc game detected, but disc 1 folder does not exist (will be checked again later)."
		return 0
	}
	
	if (-not (test-path -LiteralPath ($check_files[0].FullName + "\game.iso"))) {
		write-host "Disc 2 of Multi-Disc game detected, but game.iso doesn't exist within disc 1 folder (will be checked again later)."
		return 0
	}
	
	write-host "Disc 2 of Multi-Disc game detected: $folder_name"
	if ($check_only) { return $check_files[0].FullName }
	
	
	$disc_two_before = ($folder + "\game.iso")
	$disc_two_after = ($check_files[0].FullName + "\disc2.iso")
	
	if (-not (test-path -LiteralPath $disc_two_before)) {
		write-host "'$disc_two_before' not found to move."
		return $check_files[0].FullName
	}
	
	if (-not (test-path -LiteralPath $disc_two_after)) {
		Move-Item -LiteralPath $disc_two_before $disc_two_after
		write-host ("Moved '" + $disc_two_before + "' => '" + $disc_two_after + "'")
	} else {
		Remove-Item -LiteralPath $disc_two_before
		write-host ("Already exists: '" + $disc_two_after + "' ... Deleted: '" + $disc_two_before + "'")
	}
	Remove-Item -LiteralPath $folder
	
	return $check_files[0].FullName
}

if (test-path -LiteralPath "$PSScriptRoot\_error.txt") {
	Remove-Item -LiteralPath "$PSScriptRoot\_error.txt"
}

if (-not $md) {
	if (-not $scan -or -not (test-path $scan)) {
		write-host "Scan path '$scan' not found."
		exit
	} elseif (-not $discex -or -not (test-path -LiteralPath "$discex\DiscEx.exe")) {
		write-host "'$discex\DiscEx.exe' not found."
		exit
	} elseif (-not $gcit -or -not (test-path -LiteralPath "$gcit\gcit.exe")) {
		write-host "'$gcit\gcit.exe' not found."
		exit
	} elseif (-not $dest -or -not (test-path -LiteralPath $dest)) {
		$ret = New-Item "$dest" -type directory
		if (-not (test-path -literalpath "$dest")) {
			Write-Host "Error creating '$dest'"
			exit
		}
	}
} else {
	$ret = MultiDisc_Check -folder $dest
	exit
}

$files = Get-ChildItem -Path $scan -recurse | Where-Object { ($_.Name -match "\.(iso|gcm)$") -and ($_.FullName -notmatch "\[.{5,}\]\\(game|disc\d)\.(iso|gcm)$") } | Sort
foreach ($file in $files) {
	#if the drive doesn't exist then it might be overheating, add some sleep time
	try{
			(Get-Item "$scan").PSDrive.Name | Out-Null
			(Get-Item "$dest").PSDrive.Name | Out-Null
		}catch{
			$throttle+=20
			if($throttle/20 -gt 15){
				write-host "Drive keeps disconnecting. if it is a network drive please check your network settings. If it is a sd/usb card it might be overheating"
				exit 1
			}
		}

	write-host ""
	$ret = DoGame -iso $file.FullName
	if ($ret) { WriteError -err ("Error working on '" + $file.FullName + "'") }

	if($throttle){
		write-host "Sleeping... $throttle"
		Start-Sleep -s $throttle
	}
}

if($throttle){
	write-host "If you are using a SD/USB drive that is overheating next time use a throttle of $throttle"
}

Set-Location $PSScriptRoot

$ret = MultiDisc_Check -folder $dest
