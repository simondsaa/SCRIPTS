Function Convert-Size {
    <#
        .SYSNOPSIS
            Converts a size in bytes to its upper most value.

        .DESCRIPTION
            Converts a size in bytes to its upper most value.

        .PARAMETER Size
            The size in bytes to convert

        .NOTES
            Author: Boe Prox
            Date Created: 22AUG2012

        .EXAMPLE
        Convert-Size -Size 568956
        555 KB

        Description
        -----------
        Converts the byte value 568956 to upper most value of 555 KB

        .EXAMPLE
        Get-ChildItem  | ? {! $_.PSIsContainer} | Select -First 5 | Select Name, @{L='Size';E={$_ | Convert-Size}}
        Name                                                           Size                                                          
        ----                                                           ----                                                          
        Data1.cap                                                      14.4 MB                                                       
        Data2.cap                                                      12.5 MB                                                       
        Image.iso                                                      5.72 GB                                                       
        Index.txt                                                      23.9 KB                                                       
        SomeSite.lnk                                                   1.52 KB     
        SomeFile.ini                                                   152 bytes   

        Description
        -----------
        Used with Get-ChildItem and custom formatting with Select-Object to list the uppermost size.          
    #>
    [cmdletbinding()]
    Param (
        [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [Alias("Length")]
        [int64]$Size
    )
    Begin {
        If (-Not $ConvertSize) {
            Write-Verbose ("Creating signature from Win32API")
            $Signature =  @"
                 [DllImport("Shlwapi.dll", CharSet = CharSet.Auto)]
                 public static extern long StrFormatByteSize( long fileSize, System.Text.StringBuilder buffer, int bufferSize );
"@
            $Global:ConvertSize = Add-Type -Name SizeConverter -MemberDefinition $Signature -PassThru
        }
        Write-Verbose ("Building buffer for string")
        $stringBuilder = New-Object Text.StringBuilder 1024
    }
    Process {
        Write-Verbose ("Converting {0} to upper most size" -f $Size)
        $ConvertSize::StrFormatByteSize( $Size, $stringBuilder, $stringBuilder.Capacity ) | Out-Null
        $stringBuilder.ToString()
    }
}