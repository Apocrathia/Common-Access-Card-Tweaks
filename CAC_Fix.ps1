# Do some CAC stuff

# In case the user doesn't have it set:
# Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted

# Outlook 2007:  HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Outlook\Preferences
# Outlook 2010:  HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Outlook\Preferences
# Outlook 2013:  HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Outlook\Preferences
# Outlook 2016:  HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Preferences
# SecurityAlwaysShowButtons
# TurnOnLegacyGALdialog

$versions = "12.0", "14.0", "15.0", "16.0"
$keys = "SecurityAlwaysShowButtons", "TurnOnLegacyGALdialog"
$value = 1

foreach ($version in $versions) {
$path = "HKCU:\Software\Microsoft\Office\$version\Outlook\Preferences"
    # Test path
    if (Test-Path $path) {
        # We've got a match
        Write-Host -ForegroundColor DarkGreen "$path Found"

        foreach ($key in $keys) {
            if ((Get-ItemPropertyValue -Path $path -Name $key) -eq 1) {
                Write-Host -ForegroundColor Gray "$key value already set"
            } else {
                New-ItemProperty -Path $path -Name $key -Value $value -PropertyType DWORD
                Write-Host -ForegroundColor Green "$key value set"
            }
        }
    } else {
        # Move alone, nothing to see here
        Write-Host -ForegroundColor DarkGray "$path Not Found"
        continue
    }
}
Write-Host "Script completed successfully, please reboot Outlook!"
