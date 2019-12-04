<#
This script will connect to 365 and scan distribution groups 
that contain disabled users. If any users are found, it will
remove them. A transcript is started before it starts removing
any users from the group so that it can be reviewed after the fact. 
This file by default will be placed on your desktop.
#>

Begin {

    # Connect to Office365

    $usercredential = Get-Credential -Message 'Enter Office365 Admin Credentials:'
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    Import-PSSession $Session -DisableNameChecking -AllowClobber | Out-Null
    Connect-MsolService -Credential $usercredential

    #Get all disabled mailboxes in Office 365.

    $users = get-mailbox -Filter * | Where-Object AccountDisabled -eq True | Select-Object name

    #Gets all distribution groups in Office 365.
    
    $groups = get-distributiongroup

}

Process {

    #Start transcript

    Start-Transcript -Path $env:userprofile\desktop\DistributionGroupCleanup.txt -Force

    #Loop through each user and group; remove disabled users from distribution groups.

    foreach ($user in $users) {

        foreach ($group in $groups) {

            $members = Get-DistributionGroupMember -Identity $group.ToString()

            $members = $members.Name

            if ($members -contains $user.Name) {

                Write-Host "$($user.Name) in $group" -ForegroundColor Green
                Write-Host "Removing $($user.Name) from $group" -ForegroundColor Yellow
                Remove-DistributionGroupMember -Identity $group.Name -Member $user.Name -confirm:$false

            }
            else {

                Write-Output "$($user.Name) not in $group"

            }
        }
    }
}

End {
    
}