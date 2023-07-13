try {
    # Import the required modules
    Import-Module Microsoft.Online.SharePoint.Powershell -DisableNameChecking

    # Connect to SharePoint Online
    Connect-SPOService -Url "https://your-tenant-admin.sharepoint.com" -Credential (Get-Credential)

    # Create an array to store the site and subsite details
    $ownerAddedSites = @()

    # Get all sites within the tenant
    Get-SPOSite -Limit All | ForEach-Object {
        try {
            # Connect to the site
            Connect-PnPOnline -Url $_.Url -Credentials (Get-Credential)

            # Check if the specific account is already an owner
            $currentOwners = Get-SPOSite -Identity $_.Url | Select-Object -ExpandProperty Owner
            if ($currentOwners -notcontains "user@domain.com") {
                # Add the specific account as an owner to the current site
                Set-SPOSite -Identity $_.Url -Owner "user@domain.com"
                
                # Add site details to the ownerAddedSites array
                $ownerAddedSites += [PSCustomObject]@{
                    SiteUrl = $_.Url
                    OwnerAdded = "Yes"
                }
            } else {
                # Add site details to the ownerAddedSites array
                $ownerAddedSites += [PSCustomObject]@{
                    SiteUrl = $_.Url
                    OwnerAdded = "No"
                }
            }

            # Get all subsites within the current site
            Get-SPOSite -Limit All -Filter "ParentUrl -eq '$($_.Url)'" | ForEach-Object {
                try {
                    # Check if the specific account is already an owner
                    $subsiteCurrentOwners = Get-SPOSite -Identity $_.Url | Select-Object -ExpandProperty Owner
                    if ($subsiteCurrentOwners -notcontains "user@domain.com") {
                        # Add the specific account as an owner to the current subsite
                        Set-SPOSite -Identity $_.Url -Owner "user@domain.com"
                        
                        # Add subsite details to the ownerAddedSites array
                        $ownerAddedSites += [PSCustomObject]@{
                            SiteUrl = $_.Url
                            OwnerAdded = "Yes"
                        }
                    } else {
                        # Add subsite details to the ownerAddedSites array
                        $ownerAddedSites += [PSCustomObject]@{
                            SiteUrl = $_.Url
                            OwnerAdded = "No"
                        }
                    }
                }
                catch {
                    Write-Host "Error adding owner to subsite: $($_.Url) - $($_.Exception.Message)"
                }
            }

            # Disconnect from the current site
            Disconnect-PnPOnline
        }
        catch {
            Write-Host "Error processing site: $($_.Url) - $($_.Exception.Message)"
        }
    }

    # Disconnect from SharePoint Online
    Disconnect-SPOService

    # Export the ownerAddedSites array to a CSV file
    $ownerAddedSites | Export-Csv -Path "OwnerAddedSites.csv" -NoTypeInformation
}
catch {
    Write-Host "Error executing the script: $($_.Exception.Message)"
}
