param (
	[string]$RootPath,
	[string]$Log
)


function Take-Ownership {
	param(
		[String]$Folder
	)
	takeown.exe /A /F $Folder
	$CurrentACL = Get-Acl $Folder
	write-host ...Adding ADMINISTRATORS to $Folder -Fore Yellow
	$SystemACLPermission = "BUILTIN\Administrators","FullControl","ContainerInherit,ObjectInherit","None","Allow"
	$SystemAccessRule = new-object System.Security.AccessControl.FileSystemAccessRule $SystemACLPermission
	$CurrentACL.AddAccessRule($SystemAccessRule)
#	write-host ...Adding Area52\Tyndall Base Sysadmins to $Folder -Fore Yellow
#	$AdminACLPermission = "Area52\Tyndall Base Sysadmins","FullControl","ContainerInherit,ObjectInherit","None","Allow"
#	$SystemAccessRule = new-object System.Security.AccessControl.FileSystemAccessRule $AdminACLPermission
#	$CurrentACL.AddAccessRule($SystemAccessRule)
	Set-Acl -Path $Folder -AclObject $CurrentACL
}

function Test-Folder($FolderToTest){
	$error.Clear()
	Get-ChildItem $FolderToTest -Recurse -ErrorAction SilentlyContinue | Select FullName
	if ($error) {
		foreach ($err in $error) {
			if($err.FullyQualifiedErrorId -eq "DirUnauthorizedAccessError,Microsoft.PowerShell.Commands.GetChildItemCommand") {
				Write-Host Unable to access $err.TargetObject -Fore Red
				Write-Host Attempting to take ownership of $err.TargetObject -Fore Yellow
				Take-Ownership($err.TargetObject)
				Test-Folder($err.TargetObject)
			}
		}
	}
}
Start-Transcript $Log
Take-OwnerShip ($RootPath)
Test-Folder($RootPath)
Stop-Transcript 