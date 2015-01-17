#------------------------------------------------------------------------------
#PROCESS: 		SSHUploadWithKey.ps1
#PARAMS: 		$HostName, $HostPort, $FilePathLocal, $FilePathRemote,
#				$Username, $AuthKeyPath, $Password (OPTIONAL)
#USAGE:			SSHUploadWithKey "sftp.ssh.com" "22" "C:\outgoing\*.txt"
#				"/In/" "Testuser" "C:\path\to\openssh.key"
#AUTHOR: 		Nicholas Raymond
#LAST EDITED:	16 JAN 2015
#
#DESCRIPTION:	This script connects to a remote SFTP server utilizing a
#				SSH key file for authentication and uploads the files
#				specified. Wildcards can be used in the $FilePathLocal
#				parameter.
#------------------------------------------------------------------------------
Function SSHUploadWithKey {
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
		[string]$Password = ""
	)

#	Ensure local file for upload exists
	if( -not (Test-Path $FilePathLocal)) {
		throw("Unable to locate the local file for upload (path=${FilePathLocal})")
	}
	
#	Ensure Source/Dest paths do not end with slashes
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

#	Determine if file extension is provided for local & remote file(s)
	$FileExtLocal = [System.IO.Path]::GetExtension("$FilePathLocal")
	$FileExtRemote = [System.IO.Path]::GetExtension("$FilePathRemote")
	
#	If a directory transfer, determine the file names to transfer
	if($FileExtLocal -eq "" -and $FileExtRemote -eq "") {
		$FilesToUpload = @()
		$UploadType = "Multi"
		$FilesFound = Get-ChildItem $FilePathLocal
		foreach($_.Name in $FilesFound) {
			$FilesToUpload += $_
		}
	}
#	If no remote file extension is specified, determine the local file names to transfer
	elseif($FileExtLocal -ne "" -and $FileExtRemote -eq "") {
		$FilesToUpload = @()
		$UploadType = "Multi"
		$FileDirLocal = Split-Path -Path $FilePathLocal -Parent
		$FilesFound = Get-ChildItem $FilePathLocal
		foreach($_.Name in $FilesFound) {
			$FileValues = Split-Path -Path $_ -Leaf
			$FilesToUpload += $FileValues
			} 
		}

#	If both extensions are provided, determine if the file name is known or not
	elseif ($FileExtLocal -ne "" -and $FileExtRemote -ne "") {
		$UploadType = "Single"
		$FileDirLocal = Split-Path -Path $FilePathLocal -Parent
		$FileNameLocal = Split-Path -Path $FilePathLocal -Leaf -Resolve
		$FilesToUpload = $FileNameLocal
	}

#	Check to ensure multiple files are not being uploaded to a single file path
	if($UploadType -eq "Multi" -and $FileExtRemote -ne "") {
		throw "You cannot select a single file destination when uploading multiple files (files=${FilesToUpload}, uploadpath=${FilePathRemote})"
	}
	
#	Determine if password is provided or not, connect and validate the open connection
	if($Password -ne "") {
		Write-Message "Password specified, using alternate connection (Password + Key)"
		$SFTPConnect = Open-SFTPServerWithPublicKey -serverAddress $HostName -userName $Username -userPassword $Password -publicKeyFile $AuthKeyPath -serverPort $HostPort
	}
	else {
		Write-Message "Password not specified, using default connection (Key Only)"
		$SFTPConnect = Open-SFTPServerWithPublicKey -serverAddress $HostName -userName $Username -publicKeyFile $AuthKeyPath -serverPort $HostPort
	}
	if( -not ($SFTPConnect.Connected)) {
	    throw("Unable to establish a SSH session with host (host=${HostName})")
	}
	
#	Check to ensure remote directory exists for multi-file transfers
	if($UploadType -eq "Multi") {
		$FileParentRemote = $FilePathRemote | Split-Path -Parent
		$FileParentRemote = $FileParentRemote.Replace("\","/")
		$FileDirRemote = $FilePathRemote | Split-Path -Leaf
		$CheckFilePathRemote = $SFTPConnect.GetDirList("$FileParentRemote")
		if($CheckFilePathRemote -notcontains $FileDirRemote) {
			throw("Remote path not found (path=${FilePathRemote})")
		}
	}
	
#	Check to ensure remote directory exists for single-file transfer
	if($UploadType -eq "Single" -and $FileExtLocal -ne "") {
		$FilePathRemoteNoFile = $FilePathRemote | Split-Path -Parent
		$FileParentRemote = $FilePathRemoteNoFile | Split-Path -Parent
		$FileParentRemote = $FileParentRemote.Replace("\","/")
		$FileDirRemote = $FilePathRemoteNoFile | Split-Path -Leaf
		$CheckFilePathRemote = $SFTPConnect.GetDirList("$FileParentRemote")
		if($CheckFilePathRemote -notcontains $FileDirRemote) {
			throw("Remote path not found (path=${FilePathRemoteNoFile})")
		}
	}
	
#	Upload files if there are multiple files found and no files specified
	if($UploadType -eq "Multi" -and $FilesToUpload.Count -gt 1 -and $FileExtLocal -eq "") {
		foreach($_ in $FilesToUpload) {
			$FullUploadPath = $FilePathLocal + "\" + $_
			$SFTPConnect.Put($FullUploadPath, $FilePathRemote)
		}
	}
#	Upload files if there are multiple files found and file extensions are specified
	if($UploadType -eq "Multi" -and $FilesToUpload.Count -gt 1 -and $FileExtLocal -ne "") {
		$FileDirLocal = $FilePathLocal | Split-Path -Parent
		foreach($_ in $FilesToUpload) {
			$FullUploadPath = $FileDirLocal + "\" + $_
			$SFTPConnect.Put($FullUploadPath, $FilePathRemote)
		}
	}
	
#	Upload file if a single file is found
	if($UploadType -eq "Single") {
		if($FileExtLocal.Contains("*")){
			$FileToUpload = $FileDirLocal + "\" + $FileNameLocal
			$SFTPConnect.Put($FileToUpload, $FilePathRemote)
		}	
		else {
			$SFTPConnect.Put($FilePathLocal, $FilePathRemote)
		}
	}

#	Verify multiple files were transferred successfully
	if($UploadType -eq "Multi") {
		$CheckFileUploadExists = $SFTPConnect.GetFileList("${FilePathRemote}")
		foreach($_ in $FilesToUpload) {
			if($CheckFileUploadExists -contains $_) {
				Write-Message "VALIDATION: File $_ was found in remote path ${FilePathRemote}"
			}
			if($CheckFileUploadExists -notcontains $_) {
				throw("File $_ was not found in remote path ${FilePathRemote}")
			}
		}
	}
	
#	Verify single file was transferred successfully
	if($UploadType -eq "Single") {
		$FileNameRemote = $FilePathRemote | Split-Path -Leaf
		$FilePathRemoteNoFile = $FilePathRemoteNoFile.Replace("\","/")
		$CheckFileUploadExists = $SFTPConnect.GetFileList("${FilePathRemoteNoFile}")
		if($CheckFileUploadExists -contains $FileNameRemote) {
				Write-Message "VALIDATION: File ${FileNameRemote} was found in remote path ${FilePathRemoteNoFile}"
		}
	}
	
#	Disconnect from SSH server
	$SFTPConnect.Close()
}