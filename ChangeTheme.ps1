# usage : ChangeTheme.ps1 -Image "c:\path\to\image-without-extension" -Style Dark|Light -Type PNG|JPG|JPEG -Update True|False

param (
    [parameter(Mandatory=$False)]
    # Provide path to image
    [string]$Image,
    # Provide wallpaper style that you would like applied
    [parameter(Mandatory=$False)]
    [ValidateSet('Light','Dark')]
    [string]$Style,
    # Provide update of scheduled task
    [parameter(Mandatory=$False)]
    [ValidateSet('True','False')]
    [string]$Update,
    # Provide wallpaper style that you would like applied
    [parameter(Mandatory=$True)]
    [ValidateSet('png','jpg','jpeg')]
    [string]$Type,
    # Restart process
    [parameter(Mandatory=$True)]
    [ValidateSet('True','False')]
    [string]$RestartProcess,
    # Kill process
    [parameter(Mandatory=$True)]
    [ValidateSet('True','False')]
    [string]$KillProcess
)

if ( $Update -eq "True" ) {

    # Get Sun events
    $lat = "43.6112422"
    $long = "3.8767337"
    $Daylight = (Invoke-RestMethod "https://api.sunrise-sunset.org/json?lat=$($lat)&lng=$($long)&formatted=0").results
    $Sunrise  = ($Daylight.Sunrise | Get-Date -Format "HH:mm")
    $Sunset   = ($Daylight.Sunset | Get-Date -Format "HH:mm")

    # Check if task exists
    $TaskList = "SetLightTheme","SetDarkTheme"
    
    ForEach ( $Task in $Tasklist ) {
        $GetTask = Get-ScheduledTask | Where-Object { $_.TaskName -eq "$Task" }
    
        if ( $GetTask ) {
            # update task if exist
            try {
                if ( $Task -eq "SetLightTheme" ) {
                    $U = New-ScheduledTaskTrigger -Daily -At $Sunrise
                } elseif ( $Task -eq "SetDarkTheme" ) {
                    $U = New-ScheduledTaskTrigger -Daily -At $Sunset
                }

                Set-ScheduledTask -TaskName $GetTask.TaskName -TaskPath "\ChangeTheme\" -Trigger $U >> $null
            }
            catch {
                Write-Error -Message "Unable to update Scheduled Tasks. Error was: $_" -ErrorAction Stop
            }
        }
    }
}

# Process to restart after changing theme
$ProcessList = "Outlook","WinWord","Excel","PowerPoint"
$ProcessStart = @()

Function Set-WallPaper($Image,$Style) {

    Add-Type -TypeDefinition @" 
    using System; 
    using System.Runtime.InteropServices;
  
    public class Params
    { 
        [DllImport("User32.dll",CharSet=CharSet.Unicode)] 
        public static extern int SystemParametersInfo (Int32 uAction, 
                                                       Int32 uParam, 
                                                       String lpvParam, 
                                                       Int32 fuWinIni);
    }
"@ 
  
    $SPI_SETDESKWALLPAPER = 0x0014
    $UpdateIniFile = 0x01
    $SendChangeEvent = 0x02
  
    $fWinIni = $UpdateIniFile -bor $SendChangeEvent
  
    $ret = [Params]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $Image, $fWinIni)
 
}

if ( $Style -ne $null ) {
    Set-WallPaper -Image "$Image-$Style.$Type"
} else {
    Set-WallPaper -Image "$Image.$Type"
}

# Kill process
if ( $KillProcess -eq "True" ) {
    ForEach ( $Process in $ProcessList ) {
        if ( (Get-Process -Name "$($Process)" -ErrorAction SilentlyContinue) -ne $null ) {
            $ProcessStart=$ProcessStart+"$($Process)"
            Stop-Process -Name $Process -Force
        }
    }
}

if ( $Style -eq "Dark" ) {
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value "0" -Force >> $null
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common" -Name "UI Theme" -Value "4" -Force
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Roaming\Identities\c91873e5-c732-42c6-8034-a5959d53877c_ADAL\Settings\1186\{00000000-0000-0000-0000-000000000000}" -Name "Data" -Value ([byte[]](0x04,0x00,0x00,0x00)) -Force
}

if ( $Style -eq "Light" ) {
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value "1" -Force >> $null
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common" -Name "UI Theme" -Value "0" -Force
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Roaming\Identities\c91873e5-c732-42c6-8034-a5959d53877c_ADAL\Settings\1186\{00000000-0000-0000-0000-000000000000}" -Name "Data" -Value ([byte[]](0x00,0x00,0x00,0x00)) -Force
}

# Restart process
if ( $RestartProcess -eq "True" ) {
    ForEach ( $Process in $ProcessStart ) {
        Start-Process $Process
    }
}