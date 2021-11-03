param([string[]] $InObject = @([char] 0x0))
Function GetAsciiCode ([char] $gacChar, [int] $gacCode) {
    if ($gacCode -le 0) {
        $gacAChar = [byte[]] 0
        $gacPInto = [byte[]] 0
        $gacPI437 = [byte[]] 0
    } else {
        $gacEUnic = [System.Text.Encoding]::GetEncoding(1200)
        $gacET437 = [System.Text.Encoding]::GetEncoding(437)
        $gacETarg = [System.Text.Encoding]::GetEncoding($gacCode)
        $gacAChar = $gacEUnic.GetBytes($gacChar)
        $gacPInto = [system.text.encoding]::Convert($gacEUnic,$gacETarg,$gacAChar)
        $gacPFrom = [system.text.encoding]::Convert($gacETarg,$gacEUnic,$gacPInto)
        $gacPI437 = [system.text.encoding]::Convert($gacEUnic,$gacET437,$gacAChar)
        if ( -not ( $gacChar -eq $gacEUnic.GetString($gacPFrom) -or $gacPInto -le 31 )) 
            { $gacPInto = [byte[]] 0 }
        <#
        if ($gacChar -eq '§') {
            Write-Host "abc- " -NoNewline
            Write-Host $gacCode, AChar, $gacAChar, PInto, $gacPInto, PFrom, $gacPFrom, PI437, $gacPI437 -NoNewline
            Write-Host " -def"
        }
        #>
    }
    switch ($gacPInto.Count) {
        2 { # double-byte character set (DBCS) recognized
            [int32] $gacPInNo = $gacPInto[1]+$gacPInto[0]*256
            # [int32] $gacPInNo = 0
          }
        1 { # single-byte character set (SBCS) recognized
            [int32] $gacPInNo = $gacPInto[0]
          }
        default { [int32] $gacPInNo = 0 }
    }
    Return @($gacPInNo, $gacPI437[0])
}

<#
language groups   : https://msdn.microsoft.com/en-us/goglobal/bb688174
input method (IME): Get-WinUserLanguageList
language examples : https://www.microsoft.com/resources/msdn/goglobal/default.mspx
code pages & LCIDs: [System.Globalization.CultureInfo]::GetCultures(
                    [System.Globalization.CultureTypes]::AllCultures)|
                        Format-Custom -Property DisplayName, TextInfo
#>
$KbdLayouts = @(
   # Basic Collection (installed on all languages of the OS)
    @('0409', 437, 1252, 'en-US',  1, 'US & Western Eu'),
    @('0809', 850, 1252, 'en-GB',  1, 'US & Western Eu'),
    @('0405', 852, 1250, 'cs-CZ',  2, 'Central Europe'),
    @('0425', 775, 1257, 'et-EE',  3, 'Baltic'),
    @('0408', 737, 1253, 'el-GR',  4, 'Greek'),
    @('0419', 866, 1251, 'ru-RU',  5, 'Cyrillic'),
    @('041f', 857, 1254, 'tr-TR',  6, 'Turkic'),
   # East Asian collection: double-byte character sets (DBCS): 
    #@('0411',   0,  932, 'ja-JP',  7, 'Japanese'),     # (Japan),  DBCS
    #@('0412',   0,  949, 'ko-KR',  8, 'Korean'),       # (Korea),  DBCS
    #@('0404',   0,  950, 'zh-TW',  9, 'Trad. Chinese'),# (Taiwan), DBCS
    #@('0804',   0,  936, 'zh-CN', 10, 'Simpl.Chinese'),# (China),  DBCS
   # Complex script collection (always installed on Arabic and Hebrew localized OSes)  
    @('041E',   0,  874, 'th-TH', 11, 'Thai'),         # (Thailand)
    @('040D', 862, 1255, 'he-IL', 12, 'Hebrew'),       # (Israel)
    @('0C01', 720, 1256, 'ar-EG', 13, 'Arabic'),       # (Egypt)
    @('042A',   0, 1258, 'vi-VN', 14, 'Vietnamese'),   # (Vietnam)
   # unknown supported code page
   #@('0445',   0,    0, 'bn-IN', 15, 'Indic'),        # Bengali (India)
   #@('0437',   0,    0, 'ka-GE', 16, 'Georgian'),     # (Georgia)
   #@('042B',   0,    0, 'hy-AM', 17, 'Armenian'),     # (Armenia)
    @('0000',  -1,   -1, 'xx-xx', 99, 'dummy entry'))  # (last array element - not used)
   #@(LCID, OEM-CP, ANSI-CP, IMEtxt, GroupNo, GroupTxt)
$currentLocale = Get-WinSystemLocale
$currentIME    = "{0:x4}" -f $currentLocale.KeyboardLayoutId
$currentOCP    = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage").OEMCP
$currentACP    = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage").ACP
$currentHead   = 'IME ' + $currentIME + '/' + $currentLocale.Name + 
    "; CP" + $currentOCP + "; ANSI " + $currentACP
$currHeadColor = "Cyan"
$currCharColor = "Yellow"
# write header $InObject
Write-Host $("{0,2} {1,7} {2,7} {3,12}{4,7}{5,7}" -f `
   "Ch", "Unicode", "Alt?", "CP    IME", "Alt", "Alt0") -NoNewline
Write-Host $("    {0}" -f $currentHead) -ForegroundColor $currHeadColor
[string] $sX = ''
for ($i = 0; $i -lt $InObject.Length ; $i++) {
    [char] $sAuX = [char] 0x0
    [string] $sInX = $InObject[$i]
    if ($sInX -eq '') { [string] $sInX = [char] 0x00 }
    Try {   [int] 0 + $sInX | Out-Null
            [char] $sAuX = $sInX | Invoke-Expression
        }
    Catch { [string] $sAuX = ''} #Finally {#$sInX    += $sAuX }
    if ($sAuX -eq '') { $sX += $sInX } else { $sX += $sAuX }
}

for ($i = 0; $i -lt $sX.Length ; $i++) {
   [char] $Ch = $sX.Substring($i,1)
   $ChInt = [int] $Ch
   $ChModulo = $ChInt%256
   $altPDesc = "…$ChModulo…"
   Try {    
       # Get-CharInfo module downloadable from http://poshcode.org/5234
       #        to add it into the current session: use Import-Module cmdlet
       $Ch | Get-CharInfo |% {
           $ChUCode = $_.CodePoint
           $ChCtgry = $_.Category
           $ChDescr = $_.Description
       }
   }
   Catch {
       $ChUCode = "U+{0:x4}" -f $ChInt
       if ( $ChInt -le 0x1F -or ($ChInt -ge 0x7F -and $ChInt -le 0x9F)) 
            { $ChCtgry = "Control" } else { $ChCtgry = "" }
       $ChDescr = ""
   }
   Finally { $ChOut = $Ch }
   $altPCode = "$ChInt" # possible  Alt+ code 
   $altRCode = ""       # effective Alt+ code
   $altRZero = ""       # effective Alt+0 code
   if ( $ChCtgry -eq "Control" ) { # possibly non-printable character
      $ChOut = ''       
      $altPCode = ""
      if ($ChInt -gt 0) { $altRZero = "0$ChInt" } 
   } else {
      $ChOut = $Ch                 # supposedly printable character
      if ($ChInt -le 127) {
         $altRZero = "0$ChInt"
         $altRCode = "$ChInt"
      }
   }
   Write-Host "" # for better output readability?
   Write-Host ("{0,2} {1,7} {2,7} {3,12}{4,7}{5,7}" -f `
      $ChOut, $ChUCode, $altPCode, $altPDesc, $altRCode, $altRZero) -NoNewline
   Write-Host ("    {0}" -f $ChDescr) -ForegroundColor $currCharColor
   $altRCode = ""
   if ($ChInt -gt 127) {
      for ($j = 0; $j -le ($KbdLayouts.Length -1) ; $j++) {
         $altPCode = ""
         $altRCode = ""
         $altRZero = ""
         [int] $ACP = $KbdLayouts[$j][2]    # ANSI code page
         $aaCode = GetAsciiCode $Ch $ACP
         $xxCode = $aaCode[0]
         if ($xxCode -eq 0) {} else { $altRZero = "0$xxCode" }
         [int] $OCP = $KbdLayouts[$j][1]    # OEM code page
         $ooCode = GetAsciiCode $Ch $OCP
         $yyCode = $ooCode[0]
         if ($yyCode -eq 0) { } else { $altPCode = "$yyCode" }
         if (($altPCode + $altRZero) -ne "") { # locale-dependent line
            $ChOut = ""
            $ChUCode = ""
            if ($OCP -le 0) { $altPDesc = ''   #  not valid OEM CP
            } else          { $altPDesc = ('CP' + [string]$OCP)
            }
            $altPDesc += ($KbdLayouts[$j][3].PadLeft(6))
           #if ($KbdLayouts[$j][0] -eq $currentIME -or $yyCode -le 128) {
            if ($OCP -eq [int]$currentOCP -or $yyCode -le 128) {
                if ($yyCode -eq $ooCode[1]) { $altRCode = $altPCode }
            }
            if ($ooCode[1] -ge 1 -and $ooCode[1] -le 31 -and $altRCode -eq "") {
                $altRCode = $ooCode[1]
            }
            if ($ACP -gt 0) {
                $alt0Desc = '(ANSI' + ([string]$ACP).PadLeft(5) + 
                             ') ' + $KbdLayouts[$j][5].PadRight(16)
            } else {
                $alt0Desc = ''
            }
            if ($OCP -eq [int]$currentOCP -and  $altRCode -eq "") {
                $altRCode =  $altPCode
            }
            $line = "{0,2} {1,7} {2,7} {3,12}{4,7}{5,7}    {6}" -f `
                $ChOut, $ChUCode, $altPCode, $altPDesc, $altRCode, $altRZero, $alt0Desc
            if ($OCP -eq [int]$currentOCP) {
                Write-Host $line -ForegroundColor $currHeadColor
            } else {
                Write-Host $line
            }
         } 
      }
   }
}
# write footer
Write-Host `r`n($InObject -join ",") -ForegroundColor $currCharColor
if ($sX -eq '') {                     # simple help
   $aux = $MyInvocation.InvocationName
   "Usage  : $aux [<string>]`r`n"
   "Column :  description of character base line"
   Write-Host "       : -description of locale-dependent lines" -NoNewline
   Write-Host " (coloured for system defaults)" -ForegroundColor $currHeadColor
   "-------"
   "Ch     :  a character itself if printable"
   "Unicode:  character code (Unicode notation)"
   "Alt?   :  character code (decimal) = Alt+ code if <=127 or > 255 (unicode apps)"
   "       : -Alt+ code if following CP and IME corresponds to system default OEM-CP"
   "CP     : -OEM code page corresponding to an input method"
   "IME    :  …character code modulo 256… (note surrounding ellipses)"
   "       : -keyboard layout (input method) (text)"
   "Alt    : -effective ALT+  code complying with system default OEM-CP request"
   "Alt0   : -effective ALT+0 code for an IME corresponding to ANSI-CP"
   Write-Host "IME    :  Unicode name of a character " -NoNewline
   Write-Host "(only if activated Get-CharInfo module)" -ForegroundColor $currCharColor
   "         -(ANSI codepage) Laguage group name`r`n"
   #Write-Host ""
}