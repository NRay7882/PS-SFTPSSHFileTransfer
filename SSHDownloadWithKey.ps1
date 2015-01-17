#------------------------------------------------------------------------------
#PROCESS: 		SSHDownloadWithKey.ps1
#PARAMS: 		$HostName, $HostPort, $FilePathLocal, $FilePathRemote,
#				$Username, $AuthKeyPath, $DeleteFileRemote, $Password
#				(OPTIONAL)
#USAGE:			SSHDownloadWithKey "sftp.ssh.com" "22" "C:\incoming\" "/Out/"
#				"Testuser" "C:\path\to\openssh.key"
#AUTHOR: 		Nicholas Raymond
#LAST EDITED:	16 JAN 2015
#
#DESCRIPTION:	This script connects to a remote SFTP server utilizing a
#				SSH key file for authentication and downloads the files
#				specified. Wildcards can be used in the $FilePathRemote
#				parameter.
#------------------------------------------------------------------------------
Function SSHDownloadWithKey {
	Param(
		[Parameter(Mandatory=$true)]
		[ValidateNotNull()]
		[string] $HostName = $(throw "HostName parameter is required"),
		[Parameter(Mandatory=$true)]
		[ValidateNotNull()]
		[int] $HostPort = $(throw "HostPort parameter is required"),
		[Parameter(Mandatory=$true)]
		[ValidateNotNull()]
		[string] $FilePathLocal = $(throw "FilePathLocal parameter is required"),
		[Parameter(Mandatory=$true)]
		[ValidateNotNull()]
		[string] $FilePathRemote = $(throw "FilePathRemote parameter is required"),
		[Parameter(Mandatory=$true)]
		[ValidateNotNull()]
		[string] $Username = $(throw "Username parameter is required"),
		[Parameter(Mandatory=$true)]
		[ValidateNotNull()]
		[string] $AuthKeyPath = $(throw "AuthKeyPath parameter is required"),
		[Parameter(Mandatory=$true)]
		[ValidateNotNull()]
		[string] $DeleteFileRemote = $(throw "DeleteFileRemote parameter is required"),
		[string]$Password = ""
	)
	
#	Ensure source/dest paths do not end with slashes
	if($FilePathLocal.EndsWith("\")) {
		$FilePathLocal = $FilePathLocal.TrimEnd("\")
	}

	if($FilePathRemote.EndsWith("/")) {
		$FilePathRemote = $FilePathRemote.TrimEnd("/")
	}

#	Set working path, .EXE path
	$Invocation = (Get-Variable MyInvocation -Scope 1).Value
	$SshModuleDirectory = Split-Path $Invocation.MyCommand.Path
	$ExePath = "${SshModuleDirectory}\lib"
	Set-Location $SshModuleDirectory
	
#	Import required .DLLs for operation
	Import-Module -Name "${ExePath}\KTools.PowerShell.SFTP.dll"
	Import-Module -Name "${ExePath}\Tamir.SharpSSH.dll"
	Import-Module -Name "${ExePath}\DiffieHellman.dll"
	Import-Module -Name "${ExePath}\Org.Mentalis.Security.dll"

#	Determine file extensions for source/dest paths
	$FileExtLocal = [System.IO.Path]::GetExtension("$FilePathLocal")	
	$FileExtRemote = [System.IO.Path]::GetExtension("$FilePathRemote")

#	If both extensions are provided, check FTP for existing file to download
	if($FileExtLocal -ne "" -and $FileExtRemote -ne "") {
		$FilesToDownload = ""
		$DownloadType = "Single"
		$FileDirRemote = $FilePathRemote | Split-Path -Parent
		$FileDirRemote = $FileDirRemote.Replace("\","/")
		$FileNameRemote = $FilePathRemote | Split-Path -Leaf
		Write-Message "FILECHECK: Checking for existance of ${FilePathRemote} on ${HostName}..."
		
		if($Password -ne "") {
			$SFTPConnect = Open-SFTPServerWithPublicKey -serverAddress $HostName -userName $Username -userPassword $Password -publicKeyFile $AuthKeyPath -serverPort $HostPort
		}
		else {
			$SFTPConnect = Open-SFTPServerWithPublicKey -serverAddress $HostName -userName $Username -publicKeyFile $AuthKeyPath -serverPort $HostPort
		}
		if( -not ($SFTPConnect.Connected)) {
	    	throw("Unable to establish a SSH session with host (host=${HostName})")
		}
		if($FileExtRemote.Contains("*")){
			$FilesToDownload += $SFTPConnect.GetFileList("${FilePathRemote}")
			$SFTPConnect.Close()
			if($FilesToDownload -notlike $FileNameRemote) {
				throw("Remote file not found for download (path=${FilePathRemote})")
			}
		}
		else {
			$FilesToDownload += $SFTPConnect.GetFileList("${FileDirRemote}")
			$SFTPConnect.Close()
			if( -not ($FilesToDownload.Contains($FileNameRemote))) {
				throw("Remote file not found for download (path=${FilePathRemote})")
			}
		}
	}
	
#	If only the remote extension is given, check FTP for files of those type
	if($FileExtLocal -eq "") {
		Write-Message "FILECHECK: Multiple files specified for download, fetching file list from ${HostName}..."
		$FilesToDownload = @()
		$DownloadType = "Multi"
		if($FileExtRemote -eq "") {
			$FileDirRemote = $FilePathRemote
			$FileDirRemote = $FileDirRemote.Replace("\","/")
		}
		else {
			$FileDirRemote = $FilePathRemote | Split-Path -Parent
			$FileDirRemote = $FileDirRemote.Replace("\","/")
		}
		
		if($Password -ne "") {

			$SFTPConnect = Open-SFTPServerWithPublicKey -serverAddress $HostName -userName $Username -userPassword $Password -publicKeyFile $AuthKeyPath -serverPort $HostPort
		}
		else {
			$SFTPConnect = Open-SFTPServerWithPublicKey -serverAddress $HostName -userName $Username -publicKeyFile $AuthKeyPath -serverPort $HostPort
		}
		if( -not ($SFTPConnect.Connected)) {
	    	throw("Unable to establish a SSH session with host (host=${HostName})")
		}

		$FilesToDownload += $SFTPConnect.GetFileList("${FilePathRemote}")
		$SFTPConnect.Close()
		if($FilesToDownload.Count -ge 1) {
			Write-Message "FILECHECK: File list retrieved successfully!"
			$SFTPConnect.Close()
		}
		elseif($FilesToDownload.Count -lt 1) {
			throw("Remote file not found for download (path=${FilePathRemote})")
		}
	}

#	Ensure local dir for download exists
	if($DownloadType -eq "Single") {
		$FileDirLocal = $FilePathLocal | Split-Path -Parent
		if( -not (Test-Path $FileDirLocal)) {
			throw("Unable to locate the local directory for file download (path=${FileDirLocal})")
		}
	}
	elseif($DownloadType -eq "Multi") {
		if( -not (Test-Path $FilePathLocal)) {
			throw("Unable to locate the local directory for file download (path=${FilePathLocal})")
		}
	}
	
#	Check to ensure multiple files are not being downloaded to a single file path
	if($DownloadType -eq "Multi" -and $FileExtLocal -ne "") {
		throw "You cannot download multiple files to a single file name (files=${FilesToDownload}, downloadpath=${FilePathLocal})"
	}
	
#	Determine if password is provided or not, connect and validate the open connection
	if($Password -ne "") {
		Write-Message "CONNECT: Password specified, using alternate connection (Password + Key)"
		$SFTPConnect = Open-SFTPServerWithPublicKey -serverAddress $HostName -userName $Username -userPassword $Password -publicKeyFile $AuthKeyPath -serverPort $HostPort
	}
	else {
		Write-Message "CONNECT: Password not specified, using default connection (Key Only)"
		$SFTPConnect = Open-SFTPServerWithPublicKey -serverAddress $HostName -userName $Username -publicKeyFile $AuthKeyPath -serverPort $HostPort
	}
	if( -not ($SFTPConnect.Connected)) {
	    throw("Unable to establish a SSH session with host (host=${HostName})")
	}
	
#	Download single file, delete remote file if flagged, verify transfer
	if($DownloadType -eq "Single") {
		if($FileExtRemote.Contains("*")){
			$ResolvedFile = $FileDirRemote + "/" + $FilesToDownload
			$SFTPConnect.Get($FilePathRemote, $FilePathLocal)
			if($DeleteFileRemote -eq $True) {
				Write-Message "DELETE: Delete flag used, deleting file ${FilePathRemote}..."
				$SFTPConnect.Delete($FilePathRemote)
			}
		}
		else {
			$SFTPConnect.Get($FilePathRemote, $FilePathLocal)
			if($DeleteFileRemote -eq $True) {
				Write-Message "DELETE: Delete flag used, deleting file ${FilePathRemote}..."
				$SFTPConnect.Delete($FilePathRemote)
			}
		}
		if((Test-Path $FilePathLocal) -eq $true) {
			$FileNameLocal = $FilePathLocal | Split-Path -Leaf
			Write-Message "VERIFY: File ${FileNameLocal} was successfully downloaded to ${FileDirLocal}"
		}
		else {
			throw("Unable to locate the downloaded file (path=${FilePathLocal})")
		}
	}
	
#	Download multiple files, delete remote files if flagged, verify transfer
	if($DownloadType -eq "Multi") {
		foreach($_ in $FilesToDownload) {
			$FullDownloadPath = $FileDirRemote + "/" + $_
			$FullLocalPath = $FilePathLocal + "\" + $_
			$SFTPConnect.Get($FullDownloadPath, $FullLocalPath)
			if($DeleteFileRemote -eq $True) {
				Write-Message "DELETE: Delete flag used, deleting file ${FullDownloadPath}..."
				$SFTPConnect.Delete($FullDownloadPath)
			}

			if((Test-Path $FullLocalPath) -eq $true) {
				$FileNameLocal = $FullLocalPath | Split-Path -Leaf
				Write-Message "VERIFY: File ${FileNameLocal} was successfully downloaded to ${FilePathLocal}"
			}
			else {
				throw("Unable to locate the downloaded file (path=${FilePathLocal})")
			}
		}
	}
	
#	Disconnect from SSH server
	$SFTPConnect.Close()
}