# Created by Alex Jackevics
Function Do-Wait {
param($text,$newline)

    write-host $text -nonewline
    write-host "." -nonewline
    start-sleep 1 
    write-host "." -nonewline
    start-sleep 1 
    if($newline){
    
    
        write-host "."
    
    }else{
    
        write-host "." -nonewline 
    
    }
    start-sleep 1 

}
Function Start-Timer {
param($Activity, $Seconds, $Commands)
    
    if($Seconds){    
        
        $Second = 100 / $Seconds
    
        for($w=0; $w -lt 100;){
            
                    
            Write-Progress -Activity $Activity -Status "Waiting $($Seconds) seconds to sync changes" -PercentComplete $w
            Start-Sleep -Seconds 1
            $Seconds = $Seconds - 1
            $w = $w + $Second
            
    }}
    
    if($Commands){    
        
        $Progress = 100 / $Commands.Count
        for($w=0; $w -lt 100;){
            
            Write-Progress -Activity $Activity -Status "Status: $w%" -PercentComplete $w
            Invoke-Expression $Commands[$w/$Progress]
            $w = $w + $Progress
            
        }
        
    }
    Write-Progress -Activity $Activity -Completed
}
#region connecting modules

$Credentials = Get-Credential -Message "Please enter admin credentials"
$modules = @('Connect-MsolService -Credential $Credentials'; 'Connect-AzureAD -Credential $Credentials | out-null', 'Connect-ExchangeOnline -Credential $Credentials -showbanner:$false', 'Connect-PnPOnline -Credentials $Credentials -Url "https://plexal.sharepoint.com"', 'Connect-MgGraph')
Start-Timer -Activity "Importing modules" -Commands $modules

#endregion
#region gather information from formsite
$uri = "https://fs29.formsite.com/api/v2/Qh2Ih1/forms/bsijllmipg/results"
$token = "211PdtL54IYzsowngdVal8YgB6ZGrwv6"
Cls
Do-Wait -text "Gathering all users from FormSite"
$RequestForm = Invoke-RestMethod $uri -Headers @{Authorization="Bearer $token"}

$UsersFirstName = ($RequestForm.results.items | where {$_.position -eq "1"} | Select Value).Value #get all first names from the form
$UsersPreferredName = ($RequestForm.results.items | where {$_.position -eq "2"} | Select Value).Value #get all first names from the form
$UsersLastName = ($RequestForm.results.items | where {$_.position -eq "3"} | Select Value).Value #get all last names from the form
$UsersStartDate = ($RequestForm.results.items | where {$_.position -eq "10"} | Select Value).Value #get all Start Dates from the form
Write-host "Completed" -ForegroundColor DarkGreen
#choose which user to create
for($i = 0; $i -lt $UsersFirstName.count; $i++){
    
    if($UsersPreferredName[$i]){
    
        $FirstNameToCheck = $UsersPreferredName[$i]
    
    }else{
    
        $FirstNameToCheck = $UsersFirstName[$i]
    
    }
    try{
    
        $CheckingIfExists = Get-AzureADUser -SearchString "$($FirstNameToCheck) $($UsersLastName[$i])"
        If($CheckingIfExists -eq $null){
        
            stop
        
        }

    }Catch{
        
        $CheckingIfExists = Get-AzureADUser -SearchString "$($FirstNameToCheck)$($UsersLastName[$i])"
    
    }
    
    if($CheckingIfExists){
        
        write-host "$i $($UsersFirstName[$i]) $($UsersLastName[$i]) $($UsersStartDate[$i])" -ForegroundColor DarkGreen
    
    }Else{
    
        write-host "$i $($UsersFirstName[$i]) $($UsersLastName[$i]) $($UsersStartDate[$i])" -ForegroundColor DarkRed
    
    }

}

$UserChoice = "-1"

while ([int]$UserChoice -notin 0..$($UsersFirstName.count - 1)){

    $UserChoice = Read-Host "Please select the User"
    
}
Cls
#get information for the selected user
Do-Wait -text "Gathering information about the selected user"

$UserDetails = (($RequestForm.results | where-object {$_.Items.Value -eq $UsersFirstName[$UserChoice]}).Items).Value
#Removing any additional spaces at the end if they exists for First Name,Preferred Name and Last Name
[int]$k="0"
While($k -ne "3") {

$UserDetails[$k] = $UserDetails[$k].Trim() -replace "\s+"

$k = $k+1}

$AdditionalDetails = ((($RequestForm.results | where-object {$_.Items.Value -eq $UsersFirstName[$UserChoice]}).Items).Values).Value
Write-host "Completed" -ForegroundColor DarkGreen
#confirm user has correct information from the form

if($UserDetails[1]){ 
    
        $DisplayName = "$($UserDetails[1]) $LastName"
        $GivenName = $($UserDetails[1])

}else{
    
        $DisplayName = "$FirstName $LastName"
        $GivenName = $FirstName
}


If($AdditionalDetails[0] -eq "Innovation Services"){
    
    if($UserDetails[1]){

        $FirstName=$($UserDetails[1]);
        
    }Else{
    
        $FirstName=$($UserDetails[0]);
    
    }
    $LastName=$($UserDetails[2]);
    $JobTitle=$($UserDetails[3]);
    $Department=$($AdditionalDetails[0]);
    $Elapseit=$($AdditionalDetails[1]);
    $Salesforce=$($AdditionalDetails[2]);
    $Manager=$($UserDetails[7]);
    $Office=$($AdditionalDetails[3]);
    $StartDate=$($UserDetails[9]);
    $Phone=$($AdditionalDetails[4]);
    $Laptop=$($AdditionalDetails[5]);
    $ContractType=$($AdditionalDetails[6])


}else{


    if($UserDetails[1]){

        $FirstName=$($UserDetails[1]);
        
    }Else{
    
        $FirstName=$($UserDetails[0]);
    
    }
    $LastName=$($UserDetails[2]);
    $JobTitle=$($UserDetails[3]);
    $Department=$($AdditionalDetails[0]);
    $Manager=$($UserDetails[7]);
    $Office=$($AdditionalDetails[1]);
    $StartDate=$($UserDetails[9]);
    $Phone=$($AdditionalDetails[2]);
    $Laptop=$($AdditionalDetails[3]);
    $ContractType=$($AdditionalDetails[4])

}

#removing any spaces at the end fo the string from First Name, Last Name, Job title and Manager fields
While($FirstName[-1] -eq " "){$FirstName = $FirstName.Substring(0,$FirstName.Length-1)}
While($LastName[-1] -eq " "){$LastName = $LastName.Substring(0,$LastName.Length-1)}
While($JobTitle[-1] -eq " "){$JobTitle = $JobTitle.Substring(0,$JobTitle.Length-1)}
While($Manager[-1] -eq " "){$Manager = $Manager.Substring(0,$Manager.Length-1)}


[pscustomobject]@{

    FirstName=$FirstName;
    LastName=$LastName;
    'Job Title'=$JobTitle;
    Department=$Department;
    'Salesforce License'=$Salesforce;
    'Elapseit License'=$Elapseit;
    Manager=$Manager;
    Office=$Office;
    'Start Date'=$StartDate;
    Phone=$Phone;
    Laptop=$Laptop;
    'Contract Type'=$ContractType

}
$ConfirmInformation = Read-Host "Please confirm if the information above is correct and you like to proceed! (Y = Yes, N = No)"

While(($ConfirmInformation -ne "Y") -and ($ConfirmInformation -ne "N")){

    $ConfirmInformation = Read-Host "Please confirm if the information above is correct and you like to proceed! (Y = Yes, N = No)"

}
#endregion
if($ConfirmInformation -eq "Y"){
    #region configuring pre-requisites
    $Dir = "C:\Powershell\UpdatePhoneDetails"
    $LogDir = "C:\Powershell\JML\LogStarters"
    $SharePointDir = "/Shared Documents/12. IT/Phone Numbers"
    
    $FromSecure = -join('5929333923354c753735726554503f' -split '(?<=\G.{2})',15|%{[char][int]"0x$_"}) | ConvertTo-SecureString -AsPlainText -force
    $From = -join('737570706f727440706c6578616c2e636f6d' -split '(?<=\G.{2})',18|%{[char][int]"0x$_"})
    $CC = -join('6361726c792e7361756e6465727340706c6578616c2e636f6d' -split '(?<=\G.{2})',25|%{[char][int]"0x$_"})
    $CCEmails = @($CC, $((get-user -Identity $Manager | where {$_.RecipientType -ne "MailUser" }).WindowsEmailAddress))
    $EmailCredentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $From, $FromSecure

    $body = "
        Hi, <br>
        <br>
        This is an automated email informing you that the account for <b>$("$FirstName $LastName")</b> has been created.  Please allow 1 day for their email address to be activated. <br>
        Email address: <b>$FirstName.$LastName@plexal.com</b> <br>
        <br>
        Kind regards <br>
        <br>
        <b>Plexal IT Team <br>
        For any support please raise a ticket using the email address below. <br>
        Email: support@plexal.com <br>
        <br>
        </b>
     "


    Start-Transcript -Path "$LogDir\$FirstName $LastName.log"
    #endregion
    #region add contact number
    cls
    If($Phone -eq "Yes"){
        
        While(($ConfirmPhoneNumbersExists -ne "Y") -and ($ConfirmPhoneNumbersExists -ne "N")){

            $ConfirmPhoneNumbersExists = Read-host "Has the phone number been assigned yet? (Y = Yes, N = No)"

        }
        
        If($ConfirmPhoneNumbersExists -eq "Y"){
        
            $PhoneNumber = read-host "Please type the phone number which is allocated to the user (format 07*********)"

        }Else{
        
            $PhoneNumber = "No Phone"
        
        }

    }else{
    
        $PhoneNumber = "No Phone"
    
    }
    #endregion
    #region update excel workbook PhoneNumbersAndManagers

    #Open spreadsheet with all the information from location specified
    Do-Wait -text "Adding user information to spreadsheet in $dir"
    $Excel = New-Object -ComObject Excel.Application 
    $wb = $Excel.Workbooks.Open("$Dir\PhoneNumbersandManagers.xlsx")

    #Check last available row
    $lastCellWithData = ($wb.Worksheets.Item(1)).Range("A:A").End([Microsoft.Office.Interop.Excel.XlDirection]::xlDown)
    $lastEmptyRow = ($wb.Worksheets.Item(1)).Rows($lastCellWithData.Row + 1) 
    
    # Insert Data
    
    $wb.Worksheets.Item(1).Cells.Item($lastEmptyRow.Row,1) = "$FirstName $LastName" #DisplayName
    $wb.Worksheets.Item(1).Cells.Item($lastEmptyRow.Row,2) = $Manager #Manager
    
    If($PhoneNumber -ne "No Phone"){
    
        $wb.Worksheets.Item(1).Cells.Item($lastEmptyRow.Row,3) = $("Yes") #Company Number
    
    }else{
    
        $wb.Worksheets.Item(1).Cells.Item($lastEmptyRow.Row,3) = $("") #Company Number
    
    }

    $wb.Worksheets.Item(1).Cells.Item($lastEmptyRow.Row,4) = $PhoneNumber #Phone Number
    $wb.Worksheets.Item(1).Cells.Item($lastEmptyRow.Row,5) = $JobTitle #Title
    $wb.Worksheets.Item(1).Cells.Item($lastEmptyRow.Row,6) = $Office #Locations
    $wb.Worksheets.Item(1).Cells.Item($lastEmptyRow.Row,7) = $Department #Department
    
    $rowcount=$wb.Worksheets.item(1).usedrange.rows.count
    for($i=2;$i -le $rowcount;$i++){
    
    if(!$r){

        $r = ($wb.Worksheets.Item(1).Columns.Item(1).Rows.Item($i) | where {$_.Text -eq "$FirstName $LastName"}).Row

    }}

    #Save and quit excel
    $wb.Save() 
    $excel.Quit() #Close Excel entry

    if($r){

        write-host "Completed" -ForegroundColor DarkGreen
    
    }else{
    
        write-host "Failed" -ForegroundColor DarkRed
    
    }
    #endregion
    #region create user in O365
    Do-Wait -text "Creating account for $FirstName $LastName in Azure"
    
    $PasswordProfile=New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile #create password profile
    $PasswordProfile.Password="PlexalPlexal12" #password to be set
    $PasswordProfile.ForceChangePasswordNextLogin=$False #do not require change of password when first logged in
    
    #Create user 
    New-AzureADUser -DisplayName "$FirstName $LastName" -GivenName $FirstName -SurName $LastName -UserPrincipalName "$FirstName.$LastName@plexal.com" -UsageLocation GB -MailNickName "$FirstName.$LastName" -Mobile $PhoneNumber -Department $Department -JobTitle $JobTitle -PhysicalDeliveryOfficeName $Office -PasswordProfile $PasswordProfile -AccountEnabled $true | out-null
    Start-Timer -Activity "Syncing user account" -Seconds 20
                
    try{
    
        Get-User -Identity "$FirstName.$LastName@plexal.com" -ErrorAction stop
    
    }catch{
        
        while(!(Get-User -Identity "$FirstName.$LastName@plexal.com" -ErrorAction Ignore)){
        
            Start-Timer -Activity "Waiting for user to sync to Exchange, as user is still not synced" -Seconds 30
        
        }
    
    } 

    Try{
    
        $UserObjectID = (Get-MsolUser | where { $_.DisplayName -eq "$FirstName $LastName" }).ObjectID.Guid #get ObjectID for user to be able to assign group memberships later
    
    }Catch{
    
        $UserObjectID = (Get-AzureADUser | where {$_.DisplayName -eq "$FirstName $LastName"}).ObjectID #in case there are issue gathering object ID with command above, this command will be triggered
    
    }

    If($UserObjectID){
    
        Write-Host "Completed" -ForegroundColor DarkGreen
    
    }Else{
    
        Write-Host "Failed" -ForegroundColor DarkRed

    }
   
    #endregion
    #region add group membreships on O365

    #assign groups according to department

    $import = import-csv -path "C:\powershell\jml\csv\DepartmentGroups.csv"
    
    $DepartmentGroupObject = $import.$Department.where{$_ -notlike $null}

    #If department is Operations, give option to add Facilities or Community Team group        
    if($ContractType -eq "Contractor"){
    
        $DepartmentGroupObject.remove("Plexal Team") | out-null
    
    }
    If($Department -eq "Operations"){
            
        $AdditionalGroups = $import.Additional.where{$_ -notlike $null}
        
        for($i=0; $i -lt $AdditionalGroups.Count; $i++){
        
            Write-Host $i $AdditionalGroups[$i]
        
        }

    
    $GroupChoice = Read-host "Which additional group would you like to add?"
    $DepartmentGroupObject += $($AdditionalGroups[$GroupChoice]) #add additional group to main variable
                        
    }
    

    #adding group for office location
    $DepartmentGroupObject += "Plexal Staff $Office"

    #adding Windows users group if needed
    If($Laptop -eq "Surface Pro"){
            
        $DepartmentGroupObject += "Intune Windows Users"
        
    }elseIf($Laptop -eq "Macbook"){
    
        $DepartmentGroupObject += "Intune Mac Users"
    
    }

    #adding iOS users group to a user if needed
    If($Phone -eq "Yes"){
            
        $DepartmentGroupObject += "Intune iPhone Users"
        
    }

    Write-Host "Adding $FirstName $LastName to groups:"

    #Add user to the appropriate groups
    foreach($g in $DepartmentGroupObject){

        Write-Host "Adding $g group " -NoNewline
        try{

            Add-AzureADGroupMember -ObjectId (Get-AzureADGroup | where {$_.DisplayName -eq $g}).ObjectID -RefObjectId $UserObjectID -ErrorAction SilentlyContinue | Out-Null
        
        }Catch{
        
            Add-DistributionGroupMember -Identity $g -Member $UserObjectID | Out-Null
        
        }

        $GroupsAssigned = (Get-AzureADUserMembership -ObjectId $UserObjectId).DisplayName
        
        if($GroupsAssigned -notcontains $g){
            
            for($p=0; $p -lt 5;$p++){
            
                Start-Timer -Activity "Group is still syncing (Attempt $($p+1))" -Seconds 10
                $GroupsAssigned = (Get-AzureADUserMembership -ObjectId $UserObjectId).DisplayName
            
                if($GroupsAssigned -contains $g){
            
                    $p=5
            
                }

            }
        
        }

        If($GroupsAssigned -contains $g){
        
            Write-Host "Completed" -ForegroundColor DarkGreen
        
        }Else{
        
            Write-Host "Failed" -ForegroundColor DarkRed
        
        }

    }
    #endregion
    #region assigning manager
    Do-wait -text "Assigning manager"
    
    Set-user -Identity $UserObjectID -Manager (get-user -Identity $Manager | where {$_.RecipientType -ne "MailUser" }).UserPrincipalName #add manager    
    
    if((get-user -Identity $UserObjectID | Select Manager).Manager){
    
        Write-Host "Completed" -ForegroundColor DarkGreen
    
    }Else{
    
        write-host "Failed" -ForegroundColor DarkRed
    
    }

    #endregion
    #region assign license in O365

    #Check if licenses is available to be assigned
    Do-Wait -text "Checking Licenses" -newline:$true
       
    $ImportLicenses = Import-Csv -Path "C:\Powershell\JML\CSV\Licenses.csv"
    
    foreach($license in $ImportLicenses.LicensesToAssign.where{$_ -notlike $null}){
    
        $TotalLicenses = (Get-MgSubscribedSku | where {($_.SkuPartNumber -eq $license)}).PrepaidUnits.Enabled
        $UsedLicenses = (Get-MgSubscribedSku | where {($_.SkuPartNumber -eq $license)}).ConsumedUnits
        $LicenseDisplayName = $ImportLicenses.where{$_.AllLicenses -eq $License}.'Display Name'

        If($UsedLicenses -lt $TotalLicenses){
        
            Write-host $LicenseDisplayName "is available!" -ForegroundColor DarkGreen

            Do-wait -text "Assigning" -newline:$true
        
            $LicenseSku = (Get-MgSubscribedSku | where {($_.SkuPartNumber -eq $license)}).SkuId
            Set-MgUserLicense -UserId $UserObjectID -AddLicenses @{SkuId = $LicenseSku} -RemoveLicenses @() | out-null

            while((Get-MgUserLicenseDetail -UserId $UserObjectID).SkuPartNumber -notcontains $license){
        
                Set-MgUserLicense -UserId $UserObjectID -AddLicenses @{SkuId = $LicenseSku} -RemoveLicenses @() | out-null #Assign License
        
            }
        
            $CheckLicenseAssigned = (Get-MgUserLicenseDetail -UserId $UserObjectID).SkuPartNumber -contains $license
       
            if($CheckLicenseAssigned -eq $true){
            
                Write-Host $LicenseDisplayName "assigned`n" -ForegroundColor DarkGreen
        
            }Else{
        
                Write-Host $LicenseDisplayName "not assigned`n" -ForegroundColor DarkRed
        
            }

        }Else{
    
            Write-host $LicenseDisplayName "is not available, please re-order!`n" -ForegroundColor DarkRed
    
        }
    }
    
    #endregion
    #region update vcf files
    Do-wait -text "Quering old VCF file"
    $Excel = New-Object -ComObject Excel.Application
    $wb = (($Excel.Workbooks.Open("$Dir\PhoneNumbersandManagers.xlsx")).Worksheets.Item(1)).SaveAs("$Dir\PhoneNumbersandManagers.csv", 6)
    $importcsv = import-csv -path "$Dir\PhoneNumbersandManagers.csv"
    $import = $importcsv.Where({ $_.DisplayName -ne "" })

    $CurrentDate = Get-Date -Format ddMMyy
    $filename = "$Dir\PlexalContactsCard.vcf"

    $oldvcffile = (get-childitem -path $Dir | Where-Object {$_.Name -like "PlexalContactsCard*"}).Name

    If($oldvcffile){
    
        Write-Host $oldvcffile "Has been moved to Archive folder" -ForegroundColor DarkGreen
        Move-Item -Path "$Dir\$oldvcffile" -Destination "$Dir\Archive\$oldvcffile" -Force
   
    }Else{
    
        Write-host "No old file found!" -ForegroundColor DarkGreen
    
    }

    Write-Host "Updating VCF file"
    foreach($user in $import){
    
        If($user.MobilePhone -ne "No Phone" -or $user.DisplayName -eq $null){

            Write-Host "Adding $($user.DisplayName) Contact"
            Add-Content -Path $filename "BEGIN:VCARD"
            Add-Content -Path $filename "VERSION:2.1"
            Add-Content -Path $filename ("N;LANGUAGE=en-us:$($user.DisplayName.Split(" ")[1]);$($user.DisplayName.Split(" ")[0]);")
            Add-Content -Path $filename ("FN: $($user.DisplayName)")
            Add-Content -Path $filename ("ORG: Plexal")
            Add-Content -Path $filename ("TITLE:$($user.title)")
            Add-Content -Path $filename ("TEL;WORK;VOICE:" + $($user.MobilePhone))
            Add-Content -Path $filename "END:VCARD"
    
        }else{
        
            Write-host "$($user.DisplayName) has no phone allocated" -ForegroundColor DarkRed
        
        }
    }
    Write-host "Completed" -ForegroundColor DarkGreen
    
    Do-wait -text "Removing temporary csv file"
    $Excel.Quit()
    Start-Timer -Activity "Waiting for Excel module to close" -Seconds 5
    
    Try{
    
        remove-item -Path "$Dir\PhoneNumbersandManagers.csv" -Force

    }Catch{

        Start-Timer -Activity "Excel module didn't close, waiting..." -Seconds 5
        remove-item -Path "$Dir\PhoneNumbersandManagers.csv" -Force
    
    }

    if(!(Test-Path "$dir\PhoneNumbersandManagers.csv")){
    
        Write-Host "Completed" -ForegroundColor DarkGreen
    
    }Else{
    
        Write-Host "File has not beed deleted, please check" -ForegroundColor DarkRed

    }
    #endregion
    #region Upload file to sharepoint
    Do-Wait -text "Uploading file to SharePoint"
    
    $LastModifiedDate = (get-pnpfile -Url "$SharePointDir/PlexalContactsCard.vcf").TimeLastModified.ToString("ddMMyy")

    Rename-PnPFile -ServerRelativeUrl  "$SharePointDir/PlexalContactsCard.vcf" -TargetFileName  "PlexalContactsCard($($LastModifiedDate)).vcf" -Force -OverwriteIfAlreadyExists
    move-pnpfile -SourceUrl "$SharePointDir/PlexalContactsCard($($LastModifiedDate)).vcf" -TargetUrl "$SharePointDir/Archive" -Force
    Start-Timer -Activity "Waiting to sync changes in SharePoint" -Seconds 5
    Add-PnPFile -Folder $SharePointDir -path "$Dir\PlexalContactsCard.vcf" | out-null

    If((get-pnpfile -Url "$SharePointDir/PlexalContactsCard.vcf").TimeLastModified.ToString("ddMMyy") -eq $CurrentDate){
    
        Write-Host "Completed" -ForegroundColor DarkGreen 

    }Else{
    
        Write-Host "Failed" -ForegroundColor DarkRed 
    
    }
    #endregion
    #region Send email confirmation to selected users

    Send-MailMessage -from $From -to "Community@plexal.com" -Cc $CCEmails -BodyAsHtml $body -Subject "$FirstName $LastName Account has been Created!" -Credential $EmailCredentials -SmtpServer "smtp.office365.com" -UseSsl:$true

    #endregion
    Stop-Transcript
    
    Write-Host "Press any key to continue..."
    $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | out-null
    Exit
}else{

    exit

}

