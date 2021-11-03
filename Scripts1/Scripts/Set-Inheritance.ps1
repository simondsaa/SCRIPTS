#requires -version 3.0
 
Function Set-Inheritance {
 
[cmdletbinding(SupportsShouldProcess)]
 
Param(
[Parameter(Position=0,Mandatory,HelpMessage="Enter the file or folder path",
ValueFromPipeline=$True,ValueFromPipelineByPropertyName)]
[ValidateNotNullOrEmpty()]
[Alias("PSPath")]
[string]$Path,
[switch]$NoInherit,
[switch]$NoPreserve,
[switch]$Passthru
)
 
BEGIN {
    
    Write-Verbose  "Starting $($MyInvocation.Mycommand)"     
 
} #begin
 
PROCESS {
    Try {
        $fitem = Get-Item -path $(Convert-Path $Path) -ErrorAction Stop  
    }
    Catch {
        Write-Warning "Failed to get $path"
        Write-Warning $_.exception.message
        #bail out
        Return
    }
    if ($fitem) {
    Write-Verbose ("Resetting inheritance on {0}" -f $fitem.fullname)
    $aclProperties = Get-Acl $fItem
 
    Write-Verbose ($aclProperties | Select * | out-string)
    	
    if ($noinherit) {
        Write-Verbose "Setting inheritance to NoInherit"
    if ($nopreserve) {
     #remove inherited access rules  
            Write-Verbose "Removing existing rules"          
     $aclProperties.SetAccessRuleProtection($True,$False)
     }
     else {
     #preserve inherited access rules
     $aclProperties.SetAccessRuleProtection($True,$True)
     }
    }
    else {
     #the second parameter is required but actually ignored
        #in this scenario
     $aclProperties.SetAccessRuleProtection($False,$True)
    }
    Write-Verbose "Setting the new ACL"
    #hashtable of parameters to splat to Set-ACL
    $setParams = @{
        Path = $fitem
        AclObject = $aclProperties
        Passthru = $Passthru
    }
    
    Set-Acl @setparams
    } #if $fitem
 
} #process
 
END {
 
    Write-Verbose  "Ending $($MyInvocation.Mycommand)"     
 
} #end
 
} #end function
 
Set-Alias -name sin -value Set-Inheritance

