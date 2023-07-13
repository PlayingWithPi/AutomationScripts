Import-Module Microsoft.Online.SharePoint.Powershell -DisableNameChecking

# Connect to SharePoint Online
try {
    # Change this to fit need.
    Connect-SPOService -Url https://<SHAREPOINT-SITE.sharepoint.com> -ErrorAction Stop
} catch {
    Write-Host "Failed to connect to SharePoint Online. Please check the URL and ensure you have the necessary permissions."
    Write-Host "Error details: $_"
    exit
}

# Get all site collections
try {
    $sites = Get-SPOSite -Limit All -ErrorAction Stop | Select-Object Url
} catch {
    Write-Host "Failed to retrieve site collections. Please check your internet connection and try again."
    Write-Host "Error details: $_"
    exit
}

# Create an ArrayList to store the results
$results = New-Object System.Collections.ArrayList

# Define the exclusion list using regex and precompile them
# This will check names that "start with"
# If it doesn't block/hide "svc-", try: "svc\-" or just "svc"
$exclusionList = @("^foo", "^svc-", "^bar") | ForEach-Object { [regex]::new(^$_) }

# This checks if it contains, toggle as needed: "ForEach-Object { [regex]::new($_) }"

# Loop through each site collection
foreach ($site in $sites) {
    Write-Host "Site collection: $($site.Url)"

    # Get all site members
    try {
        $members = Get-SPOUser -Site $site.Url -ErrorAction Stop | Select-Object LoginName, DisplayName, Email
    } catch {
        Write-Host "Failed to retrieve site members for $($site.Url). Please check the site URL and your permissions."
        Write-Host "Error details: $_"
        continue
    }

    # Add site members to results array
    foreach ($member in $members) {
        # Check if user's login name is not in the exclusion list and ignore letter casing
        $excluded = $false
        foreach ($regex in $exclusionList) {
            if ($member.LoginName -cmatch $regex) {
                $excluded = $true
                break
            }
        }
        if (!$excluded) {
            $result = [ordered]@{
                "Site Collection" = $site.Url
                "User Login Name" = $member.LoginName
                "User Display Name" = $member.DisplayName
                "User Email" = $member.Email
            }
            $null = $results.Add((New-Object -TypeName PSObject -Property $result))
        }
    }
}

# Export the results to a CSV file
try {
    $results | Export-Csv -Path "<SPECIFY PATH>" -NoTypeInformation
    Write-Host "Export completed successfully. Results saved to <SPECIFY PATH>"
} catch {
    Write-Host "Failed to export results to CSV. Please check the output file path and ensure you have write permissions."
    Write-Host "Error details: $_"
}
