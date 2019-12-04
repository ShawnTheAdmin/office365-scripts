function add-officelicense {
    <#
 
    .SYNOPSIS
    Assigns an Office365 license to the specified user account.
 
    .EXAMPLE
    ./add-officelicense -username usertest -e2
 
    #>
 
    [CmdletBinding()]
    param(
 
        [Parameter(Mandatory = $true)]
        [string]$username,
 
        [switch]$e2,
 
        [switch]$e3
 
    )
 
    process {
 
        # Make sure the user exist
 
        try {
            Get-ADUser -Identity $username
        }
        catch {
            "The specified username does not exist."
            return
        }
 
        # Connect Office365
 
        Write-Verbose -Message "Connecting to Office365."
 
        if ((Get-PSSession).ConfigurationName -like "Microsoft.Exchange") {
            Write-Output "Already connected to Office365."
        }
        else {
            Write-Output "Connecting to Office365."
            $usercredential = Get-Credential -Message 'Enter Office365 Admin Credentials:'
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
            Import-PSSession $Session -DisableNameChecking -AllowClobber | Out-Null
            Connect-MsolService -Credential $usercredential
        }
 
        # Ensure there are enough of the specified licenses to proceed with script
 
       
        Write-Verbose -Message "Checking license quantity."
 
        if ($PSBoundParameters.ContainsKey('e2')) {
 
            $remaining = Get-MsolAccountSku | Where-Object { $_.AccountSkuId -like "*StandardWOff*" } | Select-Object AccountSkuID, @{n = 'RemainingUnits'; e = { $_.ActiveUnits - $_.ConsumedUnits } }
 
            if ($remaining.RemainingUnits -eq 0) {
 
                Write-Output "There are no e2 licenses remaining - canceling script."; return
            }
            else {
 
                Write-Output "There are $($remaining.RemainingUnits) e2 licenses remaining - proceeding."
 
            }
        }
 
        elseif ($PSBoundParameters.ContainsKey('e3')) {
 
            $remaining = Get-MsolAccountSku | Where-Object { $_.AccountSkuId -like "*ENTERPRISE*" } | Select-Object AccountSkuID, @{n = 'RemainingUnits'; e = { $_.ActiveUnits - $_.ConsumedUnits } }
 
            if ($remaining.RemainingUnits -eq 0) {
 
                Write-Output "There are no e3 licenses remaining - canceling script."; return
            }
            else {
 
                Write-Output "There are $($remaining.RemainingUnits) e2 licenses remaining - proceeding."
 
            }
 
        }
 
        # Ensure the user doesn't already have a license
 
       
        Write-Verbose -Message "Ensuring user isn't already licensed."
 
        if ((Get-MsolUser -UserPrincipalName "$($username)@domain.com").IsLicensed) {
 
            Write-Output "User is already licensed."; return
        }
 
        else { }
 
        # Assign specified Office365 license
 
       
        Write-Verbose -Message "Assigning license to $($username)."
 
        If ($PSBoundParameters.ContainsKey('e2')) {
 
            Set-MsolUserLicense -UserPrincipalName "$($username)@domain.com" -AddLicenses COMPANY:STANDARDWOFFPACK
 
        }
        elseif ($PSBoundParameters.ContainsKey('e3')) {
 
            Set-MsolUserLicense -UserPrincipalName "$($username)@domain.com" -AddLicenses COMPANY:ENTERPRISEPACK
 
        }
    }
}