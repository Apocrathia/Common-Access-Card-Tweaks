# Do some CAC stuff

# In case the user doesn't have it set:
# Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted

# Outlook 2007:  HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Outlook\
# Outlook 2010:  HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Outlook\
# Outlook 2013:  HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Outlook\
# Outlook 2016:  HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\

# Subpath: Preferences
#   SecurityAlwaysShowButtons
#   TurnOnLegacyGALdialog

# Subpath: Security
#   SupressNameChecks

$versions = "12.0", "14.0", "15.0", "16.0"
$keys = "SecurityAlwaysShowButtons", "TurnOnLegacyGALdialog"
$value = 1
$modified = 0

foreach ($version in $versions) {
$path = "HKCU:\Software\Microsoft\Office\$version\Outlook\Preferences"
    # Test path
    if (Test-Path $path) {
        # We've got a match
        Write-Host -ForegroundColor DarkGreen "$path Found"

        foreach ($key in $keys) {
            # Test if key exists
            try {
                if ((Get-ItemPropertyValue -Path $path -Name $key) -eq $value) {
                    Write-Host -ForegroundColor Gray "$key value already set"
                } else {
                    # Set key value
                    New-ItemProperty -Path $path -Name $key -Value $value -PropertyType DWORD
                    Write-Host -ForegroundColor Green "$key value set"
                    $modified = 1
                }
            # It doesn't
            } catch {
                # Set key value
                New-ItemProperty -Path $path -Name $key -Value $value -PropertyType DWORD
                Write-Host -ForegroundColor Green "$key value set"
                $modified = 1
            }
        }
        # Also write SupressNameChecks
        $path = "HKCU:\Software\Microsoft\Office\$version\Outlook\Security"
        $key = "SupressNameChecks"
        # Test if key exists
        try {
            if ((Get-ItemPropertyValue -Path $path -Name $key) -eq $value) {
                Write-Host -ForegroundColor Gray "$key value already set"
            } else {
                New-ItemProperty -Path $path -Name $key -Value $value -PropertyType DWORD
                Write-Host -ForegroundColor Green "$key value set"
                $modified = 1
            }
        # It doesn't
        } catch {
                New-ItemProperty -Path $path -Name $key -Value $value -PropertyType DWORD
                Write-Host -ForegroundColor Green "$key value set"
                $modified = 1
        }

    } else {
        # Move along, nothing to see here
        Write-Host -ForegroundColor DarkGray "$path Not Found"
        continue
    }
}

# Display completion text to user
if ($modified -eq 0) {
    Write-Host "Script completed successfully, no changes were made to your system."
} else {
    Write-Host "Script completed successfully, please reboot Outlook!"
}
