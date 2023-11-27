# useage: script.ps1 [updatefiles...]
$7zexe = 'C:\Program Files\7-Zip\7z.exe'

if (!(New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
	try {
		(New-Object System.IO.FileStream($PSCommandPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite)).Dispose()
	}
	catch {
		Start-Process (Get-WmiObject Win32_Process -Filter "ProcessId = '$PID'").ExecutablePath "-f `"$PSCommandPath`" $args" -Verb RunAs | Wait-Process
		exit
	}
}
$ahkrelease = Invoke-RestMethod 'https://api.github.com/repos/thqby/AutoHotkey_H/releases/latest'
if ($ahkrelease) {
	foreach ($asset in $ahkrelease.assets) {
		if ($asset.name.ToLower().EndsWith('.7z')) {
			if ($asset.updated_at.GetType().Name -eq 'String') {
				$asset.updated_at = [System.DateTime]::Parse($asset.updated_at)
			}
			$asset.updated_at = $asset.updated_at.ToLocalTime()
			Break
		}
	}
	$ahkneedupdate = 0
	$allahkfiles = 'AutoHotkey64.exe', 'AutoHotkey32.exe', 'AutoHotkey64.dll', 'AutoHotkey32.dll'
	if ($args.Length) {
		$needahkfiles = $args
	}
	else {
		$needahkfiles = $allahkfiles
	}
	foreach ($ahkfile in $needahkfiles) {
		$ahkfile = Get-Item ($PSScriptRoot + '\' + $ahkfile) -ErrorAction SilentlyContinue
		if (!$ahkfile -or ($ahkfile.CreationTime -lt $asset.updated_at)) {
			$ahkneedupdate = 1
			Break
		}
	}
	if ($ahkneedupdate) {
		$ahkfile = $PSScriptRoot + '\ahk.7z'
		try {
			Invoke-WebRequest $asset.browser_download_url -OutFile $ahkfile
			Write-Output ('The current version: ' + $ahkrelease.tag_name)
		}
		catch {
			Write-Error 'Download fail'
		}
		if (!(Get-Item $7zexe)) {
			$7zexe = where.exe 7z.exe
			if (!$7zexe) {
				$7zexe = where.exe /R 'C:\Program Files' 7z.exe
				if ($7zexe -and ($7zexe.GetType().Name -ne 'String')) {
					$7zexe = $7zexe[0]
				}
			}
		}
		if ($7zexe) {
			& $7zexe x -aoa ('-o' + $PSScriptRoot) $ahkfile $needahkfiles
			Remove-Item ($PSScriptRoot + '\ahk.7z')
			foreach ($ahkfile in $needahkfiles) {
				$ahkfile = Get-Item ($PSScriptRoot + '\' + $ahkfile)
				if ($ahkfile) {
					$ahkfile.CreationTime = $asset.updated_at
					$ahkfile.LastWriteTime = $asset.updated_at
				}
			}
		}
		else {
			Write-Error 'Cannot find 7z.exe'
		}
	}
	else {
		Write-Output "It's the latest version"
	}
}

Remove-Variable 7zexe
Remove-Variable ahkrelease
Remove-Variable asset
Remove-Variable ahkneedupdate
Remove-Variable allahkfiles
Remove-Variable needahkfiles
Remove-Variable ahkfile