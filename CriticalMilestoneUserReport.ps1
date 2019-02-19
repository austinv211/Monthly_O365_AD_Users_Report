<#
NAME: CriticalMilestoneUserReport.ps1
DESCRIPTION: Get licensed users for O365 based on business group
PREREQS: PSExcel module
AUTHOR: Austin Vargason
DATEMODIFIED: 08/03/18
#>

#TODO: Add scheduled task to run report on demand with a date/time trigger.
#TODO: Backup code online


#import the JoinObject module
import-module JoinObject

#connect to msolonline
#Connect-MsolService

#get the current date
$date = get-date -Format "MM_dd_yy"

#filepath to report
$filePath = "Monthly0365UserReport_$date.xlsx"

$lowOrgFull = Import-Excel -Path .\LowOrgReference.xlsx

#function to get joined AD and O365 data
function Get-JoinedADMsolData() {

    #o365 data
    $data = Get-MsolUser -All | 
    Select-Object UserPrincipalName, DisplayName,`
    Department, isLicensed, Licenses
    #extensionattribute3, Company
    #attributes commented out for later use

    #AD data
    $AdData = Get-ADUser -Filter "*" -Properties UserPrincipalName, DisplayName,`
        EmployeeID, Description, Department, Company, Created, msexchextensionattribute18, extensionAttribute3, mailNickname,`
        DistinguishedName, SamAccountName, LastLogonTimestamp |
                Select-Object UserPrincipalName,`
                EmployeeID,`
                @{Name="Description";Expression={$_.Description.Substring(0,20)}},`
                @{Name="AdDisplayName";Expression={$_.DisplayName}},`
                @{Name="AdDepartment";Expression={$_.Department}},`
                @{Name="LastLogonAD";Expression={[datetime]::FromFileTime($_.LastLogonTimestamp).ToString('MM-dd-yyyy')}},`
                Company,`
                Created,`
                extensionAttribute3,`
                msexchextensionattribute18,`
                mailNickname,`
                DistinguishedName,`
                SamAccountName

    #Join the Data
    $result = Join-Object -Left $data -Right $AdData -LeftJoinProperty UserPrincipalName -RightJoinProperty UserPrincipalName -Type AllInLeft

    #write the output
    Write-Output -InputObject $result
}

function Get-LowOrgHash() {

    #create a new hash for our loworg data
    $lowOrgHash = @{}

    #get a list of Loworg id and organization name
    $lowOrgList = $lowOrgFull | Select-Object LOWORG_ID, EXP_ORGANIZATION_NAME

    #for each item in the list add to the hash after trimming
    foreach ($item in $lowOrgList) {
        #set LOWORG_ID to the key and trim it
        $key = $item | Select-Object -ExpandProperty LOWORG_ID | Out-String
        $key = $key.Trim()

        #set EXP_ORGANIZATION_NAME to the value and trim it
        $value = $item | Select-Object -ExpandProperty EXP_ORGANIZATION_NAME| Out-String
        $value = $value.Trim()

        #add the key value pair to the hash
        $lowOrgHash.Add($key, $value)

    }

    foreach ($key in $lowOrgHash.Keys) {
        Write-Output @{$key = $lowOrgHash.$key}
    }
}

function Set-ReportStandardization() {
    param (
        # Input Data
        [Parameter(Mandatory=$true, ValueFromPipeline = $true)]
        [psobject[]]
        $InputData
    )

    Process {
        foreach ($row in $InputData) {
            #if the company is community services, standardize it
            if ($row.Company -like "*Community Services*") {
                $row.Company = "Community Services"
            }

            #add the gcc license attribute to the result members
            $hasEnterprise = $false

            if ($row.Licenses.AccountSkuId -contains "sdcountycagov:ENTERPRISEPACK_GOV") {
                $hasEnterprise = $true
            }

            $row | Add-Member -Name "hasEnterprisePackLicense" -Value $hasEnterprise -MemberType NoteProperty

            #if the loworg is state level child support, then mark as false for isLicensed, since cost is covered by State
            if ($row.msexchextensionattribute18 -eq "37802" -or $row.msexchextensionattribute18 -eq "37804" -or $row.msexchextensionattribute18 -eq "37818" -or $row.msexchextensionattribute18 -eq "37824" -or $row.msexchextensionattribute18 -eq "37816" -or $row.msexchextensionattribute18 -eq "37826") {
                $row.isLicensed = $false
                $row.hasEnterprisePackLicense = $false
            }
        }
    }

}

function Set-InactiveUsers() {
    param (
        # Gets the inactive users and sets their groups accordingly
        [Parameter(Mandatory=$true, ValueFromPipeline = $true)]
        [psobject[]]
        $InputData
    )

    Process {

        #TODO: Get other LOA category from caroline msexchextension19 = LOA or PLA for AD

        foreach ($row in $InputData) {
            
            #standardize Last Logon
            if ($lastLogon -eq "12-31-1600" -or $lastLogon -eq "") {
                $row.LastLogonAD = "No Logon Date Found"
            }

            #logic conditions to check whether will be proposed
            $hasOldLogon = $false            
            $notSupportAccount = ($group -ne "Support Accounts (County-Approved ITO Use)")
            $isUser = $true
            $notTerminated = $true
            $notCreatedRecently = $true

            #determine whether the login is older than 6 months
            try {
                $hasOldLogon = ([DateTime]::Parse($row.LastLogonAD) -le (Get-Date).AddMonths(-6))
            }
            catch {
                $hasOldLogon = $false
            }

            #determine whether user was created recently
            if ($null -ne $row.Created) {
                try {
                    $createdDate = [DateTime]::Parse($row.Created)
                    $notCreatedRecently = ($createdDate -le (Get-Date).AddMonths(-6))
                }
                catch {
                    Write-Host "Error Trying to Parse Creation Date for User:" $row.UserPrincipalName -ForegroundColor Red
                    Write-Host "Date:" $row.Created -ForegroundColor Red
                }
            }
            
            # determine Non User and Terminated Accounts
            if($null -ne $row.DistinguishedName) {
                $isUser = (!$row.DistinguishedName.Contains("Non-User"))
                $notTerminated = (!$row.DistinguishedName.Contains("Terminated"))
            }

            #determine whether the account is proposed for the inactive page
            $isProposed = ($hasOldLogon -and $notSupportAccount -and $isUser -and $notTerminated -and $notCreatedRecently)

            #if the last logon date is less than or equal to 6months ago, then add to the Licenses to be Evaluated for Redployment
            if ($isProposed) {
                $row.Group = "Inactive Users"
            }

        }
    }
}

function Get-LowOrgSummary() {
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [psobject[]]
        $InputData
    )

    Begin {
        #make a summary hash for lowOrgList
        $lowOrgSummary = @{}
        $lowOrgSummary.Add("No LowOrg Entry", 0)
    }
    Process {

        foreach ($row in $InputData) {

            $lowOrgName = $row.LowOrgName

            if ($lowOrgSummary.ContainsKey($lowOrgName) -and $row.hasEnterprisePackLicense -eq $true) {
                $lowOrgSummary.$lowOrgName = $lowOrgSummary.$lowOrgName + 1
            }
            elseif ($lowOrgSummary.ContainsKey($lowOrgName) -and $row.hasEnterprisePackLicense -eq $false) {
                #do nothing
            }
            elseif (!$lowOrgSummary.ContainsKey($lowOrgName) -and $row.hasEnterprisePackLicense -eq $true) {
                $lowOrgSummary.Add($lowOrgName, 1)
            }
            elseif (!$lowOrgSummary.ContainsKey($lowOrgName) -and $row.hasEnterprisePackLicense -eq $false) {
                $lowOrgSummary.Add($lowOrgName, 0)
            }
        }
    }
    End {
        foreach ($key in $lowOrgSummary.Keys) {
            Write-Output @{$key = $lowOrgSummary.$key}
        } 
    }
}

function Set-OrgName() {
    param (
        # Input Data Property
        [Parameter(Mandatory=$true, ValueFromPipeline = $true)]
        [psobject[]]
        $InputData,
        [Parameter(Mandatory=$true)]
        $lowOrgHash
    )

    Process {

        foreach ($row in $InputData) {

            #initiate orgName
            $orgName = ""

            #get the loworg id from the row
            $lowOrgId = $row.msexchextensionattribute18

            #try to get the orgName from the reference
            try {
                $orgName = $lowOrgHash.$lowOrgId
            }
            catch {
                $orgName = "No LowOrg Entry"
            }

            #add the LowOrgName to the row
            $row | Add-Member -Name LowOrgName -Value $orgName -MemberType NoteProperty

            if ($null -eq $row.LowOrgName) {
                $row.LowOrgName = "No LowOrg Entry"
            }
        }
    }
}

function Set-MsolGroups() {
    param (
        # joined msol/AD data input
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [psobject[]]
        $JoinedData
    )

    Process {
        foreach ($row in $JoinedData) {
            
            #if the username contains .onmicrosoft.com, write portal account or vendor if aeonnexus
            if ($row.UserPrincipalName.Contains(".onmicrosoft.com")) {
                if ($row.UserPrincipalName.Contains("aeonnexus")) {
                    $row.Company = "Vendor"
                }
                else {
                    $row.Company = "O365 Portal Account"
                }
            }

            #create a group for the report company
            switch ($row.Company) {
                "IT Outsourcer" {$group = "Support Accounts (County-Approved ITO Use)"}
                "DXC Technologies" {$group = "Support Accounts (County-Approved ITO Use)"}
                "Perspecta" {$group = "Support Accounts (County-Approved ITO Use)"}
                "AT&T" {$group = "Support Accounts (County-Approved ITO Use)"}
                "DXC" {$group = "Support Accounts (County-Approved ITO Use)"}
                "Test" {$group = "Support Accounts (County-Approved ITO Use)"}
                "Vendor" {$group = "Support Accounts (County-Approved ITO Use)"}
                "O365 Portal Account" {$group = "Support Accounts (County-Approved ITO Use)"}
                "" {$group = "No Company"}
                "Other Budgetary Entity" {$group = "Other Budgetary Entities"}
                default {$group = $row.Company}
            }


            #add a group member to the row
            $row | Add-Member -Name Group -Value $group -MemberType NoteProperty
        }
    }

}

function Get-LowOrgTable() {
    param(
        [Parameter(Mandatory=$true)]
        $lowOrgSummary
    )

    #create a table for the lowOrg summary page
    $lowOrgTable = @()


    #loop through the keys in the hash table and add them into the table
    foreach ($key in $lowOrgSummary.Keys) {
        #create a new object
        $obj = New-Object -TypeName PsObject

        #add the key and the value to the object
        $obj | Add-Member -Name LowOrgName -Value $key -MemberType NoteProperty
        $obj | Add-Member -Name "Number of Licenses Utilized" -Value $lowOrgSummary.$key -MemberType NoteProperty
        $obj | Add-Member -Name "Last Month" -Value "" -MemberType NoteProperty
        $obj | Add-Member -Name "Monthly Change" -Value "" -MemberType NoteProperty
        $obj | Add-Member -Name "6 Month Average" -Value "" -MemberType NoteProperty

        #add the object to the table
        $lowOrgTable += $obj
    }

    #return the lowOrg table
    Write-Output -InputObject $lowOrgTable
}

function Get-ProductSummary() {
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [psobject[]]
        $InputData
    )

    Process {

        #TODO: Add color coding for users containing disabled or dusers

        $productLicenseSummary = @()

        #add to the product license summary
        foreach ($user in $InputData) {
            
            $products = @()
    
            #if statements to add users in the list
            if ($user.Licenses.AccountSkuId -contains "sdcountycagov:POWERBI_PRO_GOV") {
                $products += "POWER BI PRO"
            }
    
            if ($user.Licenses.AccountSkuId -contains "sdcountycagov:PROJECTONLINE_PLAN_1_GOV") {
                $products += "Project Online Premium without Project Client for Government"
            }
    
            if ($user.Licenses.AccountSkuId -contains "sdcountycagov:PROJECTESSENTIALS_GOV") {
                $products += "Project Online Essentials"
            }
    
            if ($user.Licenses.AccountSkuId -contains "sdcountycagov:MCOMEETADV_GOV") {
                $products += "Audio Conferencing for Government"
            }
    
            foreach ($product in $products) {
                $obj = New-Object -TypeName PsObject
    
                $obj | Add-Member -Name UserPrincipalName -Value $user.UserPrincipalName -MemberType NoteProperty
                $obj | Add-Member -Name DisplayName -Value $user.DisplayName -MemberType NoteProperty
                $obj | Add-Member -Name EmployeeID -Value $user.EmployeeID -MemberType NoteProperty
                $obj | Add-Member -Name Department -Value $user.Department -MemberType NoteProperty
                $obj | Add-Member -Name Product -Value $product -MemberType NoteProperty
                $obj | Add-Member -Name LastLogonAD -Value $user.LastLogonAD -MemberType NoteProperty
                $obj | Add-Member -Name Company -Value $user.Company -MemberType NoteProperty
                $obj | Add-Member -Name extensionAttribute3 -Value $user.extensionAttribute3 -MemberType NoteProperty
                $obj | Add-Member -Name msexchextensionattribute18 -Value $user.msexchextensionattribute18 -MemberType NoteProperty
                $obj | Add-Member -Name mailNickname -Value $user.mailNickname -MemberType NoteProperty
                $obj | Add-Member -Name DistinguishedName -Value $user.DistinguishedName -MemberType NoteProperty
                $obj | Add-Member -Name SamAccountName -Value $user.SamAccountName -MemberType NoteProperty
                $obj | Add-Member -Name Group -Value $user.Group -MemberType NoteProperty
                $obj | Add-Member -Name LowOrgName -Value $user.LowOrgName -MemberType NoteProperty
    
                $productLicenseSummary += $obj
            }
        }
    }
    End {
        Write-Output -InputObject $productLicenseSummary
    }
}

#function to build o365 user report
function Get-MonthlyMsolReport
{


    #connect to Office 365
    Connect-MsolService

    #write to the console
    Write-Host "Connected to O365" -ForegroundColor Green

    #write to the console
    Write-Host "Exporting Result..." -ForegroundColor Yellow

    #get all users and department
    $result = Get-JoinedADMsolData

    #standardize the report
    $result | Set-ReportStandardization

    #get the loworg reference data
    $lowOrgHash = Get-LowOrgHash

    #set the orgNames
    $result | Set-OrgName -lowOrgHash $lowOrgHash

    # make a summary hash for lowOrgList
    $lowOrgSummary = $result | Get-LowOrgSummary

    # get the groups property
    $result | Set-MsolGroups

    #Get the inactive users
    $result | Set-InactiveUsers



    #create a table for the lowOrg summary page
    $lowOrgTable = Get-LowOrgTable -lowOrgSummary $lowOrgSummary

    #export the lowOrg table to the averages
    $lowOrgResult = Get-SixMonthAverages -lowOrgTable $lowOrgTable

    #get the licenses to be evaluated page
    $licensesEval = $result | Where-Object {$_.Group -eq "Inactive Users" -and $_.hasEnterprisePackLicense -eq $true}

    #add Data to the Licenses Eval Page
    foreach ($row in $licensesEval) {

        #check if the user is LOA
        $isLOA = $false

        if ($row.Description -like "*LOA*") {
            $isLOA = $true
        }

        $row | Add-Member -Name "isLOA" -Value $isLOA -MemberType NoteProperty 
    }

    #get the users for the support account page
    $supportAccountData = $result | Where-Object {$_.Group -eq "Support Accounts (County-Approved ITO Use)" -and $_.hasEnterprisePackLicense -eq $true}

    #get the monthly billing results
    $productUsers = $result | Where-Object {$_.Licenses.AccountSkuId -contains "sdcountycagov:POWERBI_PRO_GOV" -or $_.Licenses.AccountSkuId -contains "sdcountycagov:PROJECTONLINE_PLAN_1_GOV" -or $_.Licenses.AccountSkuId -contains "sdcountycagov:MCOMEETADV_GOV" -or $_.Licenses.AccountSkuId -contains "sdcountycagov:PROJECTESSENTIALS_GOV"}

    #get the product summary
    $productLicenseSummary = Get-ProductSummary -InputData $productUsers

    # if a file already exists with the file name, delete it
    if (Get-ChildItem -Path $filePath -ErrorAction SilentlyContinue)
    {
        Get-ChildItem -Path $filePath -ErrorAction SilentlyContinue|
        Remove-Item -Force
    }

    #create a pivot table definition
    $ptd = [ordered]@{}
    $ptd += New-PivotTableDefinition -PivotTableName "Sum. of Deployed O365 Licenses" -IncludePivotChart -ChartType Pie -PivotData @{'Group'='count'} -PivotRows Group -SourceWorkSheet "Licensed Users"
    $ptd += New-PivotTableDefinition -PivotTableName "Sum. of Deployment - Bus. Group" -SourceWorkSheet "Licensed Users" -PivotRows Group, Company, LowOrgName, UserPrincipalName -PivotData @{'Group'='count'}

    #export the result to an excel file, into a table with multiple worksheets
    $excel = $productLicenseSummary | Select-Object UserPrincipalName, DisplayName, EmployeeID, Department, Product, Group, LowOrgName, DistinguishedName | 
        Export-Excel -Path $filePath -TableName "BillableSubsTable" -AutoSize -WorkSheetname "Det. of Billable Subscriptions"

    $licensesEval | Select-Object UserPrincipalName, DisplayName, Department, EmployeeID, Description, LastLogonAD, Company, extensionAttribute3, msexchextensionattribute18, mailNickname, DistinguishedName, SamAccountName, @{Name="Has an O365 Entitlement/License";Expression={$_.hasEnterprisePackLicense}}, Group, LowOrgName, isLOA | 
        Export-Excel -Path $filePath -TableName "EvaluationTable" -AutoSize -WorkSheetname "Inactive Users"

    $result | Select-Object UserPrincipalName, DisplayName, Department, LastLogonAD, extensionAttribute3, msexchextensionattribute18, mailNickname, DistinguishedName, SamAccountName, @{Name="Has an O365 Entitlement/License";Expression={$_.hasEnterprisePackLicense}}, Group, LowOrgName |
        Export-Excel -Path $filePath -WorkSheetname "Det. of Deployed O365 Licenses" -AutoSize -TableName "UserDataTable"

    $result | Select-Object -Property * -ExcludeProperty Licenses, Description, EmployeeID | Where {$_.hasEnterprisePackLicense -eq $true} |
        Export-Excel -Path $filePath -WorkSheetname "Licensed Users" -AutoSize -TableName "LicensedDataTable" -PivotTableDefinition $ptd

    $ls = $lowOrgResult | Export-Excel -Path $filePath -WorkSheetname "Sum. of Deployments - LowOrg" -AutoSize -TableName "LowOrgSummaryTable" -PassThru

    #add total to lowOrg summary
    $ls.Workbook.Worksheets["Sum. of Deployments - LowOrg"].Cells["A405"].value = "Grand Total"
    
    $ls.Workbook.Worksheets["Sum. of Deployments - LowOrg"].Cells["B405"].formula = "=SUM(B1:B404)"

    $ls.Workbook.Worksheets["Sum. of Deployments - LowOrg"].Cells["C405"].formula = "=SUM(C1:C404)"

    $ls.Workbook.Worksheets["Sum. of Deployments - LowOrg"].Cells["D405"].formula = "=SUM(D1:D404)"

    $ls.Workbook.Worksheets["Sum. of Deployments - LowOrg"].Cells["E405"].formula = "=SUM(E1:E404)"

    #get conditional text for highlighting
    Add-ConditionalFormatting -Address $ls.Workbook.Worksheets["Det. of Billable Subscriptions"].Names["DistinguishedName"].Address -RuleType ContainsText -ConditionValue "disabled" -ForegroundColor Red -Bold

    $ls.Save()

    $ls.Dispose()

    #export support Account data
    $supportAccountData | Select -Property * -ExcludeProperty EmployeeID | Export-Excel -Path $filePath -TableName "SupportAccountTable" -WorkSheetname "Support Account Details" -AutoSize

    #export low org reference page
    $lowOrgFull | Export-Excel -Path $filePath -TableName "ReferenceTable" -WorkSheetname "LowOrg Reference" -AutoSize
   

}

function Get-SixMonthAverages() {

    param (
        [Parameter(Mandatory=$true)]
        [PsObject[]]$lowOrgTable
    )

    #get the list of files to calculate averages
    $files = Get-ChildItem -Path ./averages/

    #get the fileName of the files in the folder
    $fileNames = $files | Select -ExpandProperty Name

    #get the current month
    $curMonth = (get-date).Month

    #create an array of months
    $months = @("Jan", "Feb", "March", "April", "May", "June", "July", "Aug", "Sep", "Oct", "Nov", "Dec" )

    #switch to get the correct index for current month
    switch ($curMonth)
    {
        1 { $curMonth = $months[0]; $lastMonth = $months[11]}
        2 { $curMonth = $months[1]; $lastMonth = $months[0]}
        3 { $curMonth = $months[2]; $lastMonth = $months[1] }
        4 { $curMonth = $months[3]; $lastMonth = $months[2]}
        5 { $curMonth = $months[4]; $lastMonth = $months[3]}
        6 { $curMonth = $months[5]; $lastMonth = $months[4]}
        7 { $curMonth = $months[6]; $lastMonth = $months[5]}
        8 { $curMonth = $months[7]; $lastMonth = $months[6]}
        9 { $curMonth = $months[8]; $lastMonth = $months[7]}
        10 { $curMonth = $months[9]; $lastMonth = $months[8]}
        11 { $curMonth = $months[10]; $lastMonth = $months[9]}
        12 { $curMonth = $months[11]; $lastMonth = $months[10]}
    }

    #export the lowOrgTable to the current month csv
    $lowOrgTable | Export-Csv -Path ./averages/$curMonth.csv -NoTypeInformation

    #content array 
    $content = @()

    #fill the content array
    foreach ($csv in $fileNames) {

        $obj = new-object -TypeName psobject

        $data = import-csv -Path ./averages/$csv

        $obj | Add-Member -Name "MonthFile" -Value $csv -MemberType NoteProperty
        $obj | Add-Member -Name "File" -Value $data -MemberType NoteProperty

        $content += $Obj
    }

    #get the newest sheet and the sheets previous
    if ($null -eq ($content | Where-Object {$_.MonthFile -like "$curMonth*"})) {
        New-Item -ItemType File -Path .\averages -Name "$curMonth.csv"
    }

    $newestSheet = $content | Where-Object {$_.MonthFile -like "$curMonth*"}
    $previousSheets = $content | Where-Object {$_.MonthFile -ne $newestSheet.MonthFile}


    #go through the rows in the current file to calculate the average based on the previous file
    foreach ($row in $newestSheet.File) {
        $lowOrg = $row.LowOrgName

        $avg = 0
        $counter = 0

        foreach ($sheet in $previousSheets) {
            #get the value for last month
            if ($sheet.MonthFile -like "$lastMonth*") {
                $row."Last Month" = $sheet.File | Where-Object { $_.LowOrgName -eq $lowOrg} | Select-Object -ExpandProperty "Number of Licenses Utilized"
            }

            #get the 6 month averages for that sheet
            if ( $sheet.File | Where { $_.LowOrgName -eq $lowOrg} ) {
                $avg += $sheet.File | Where {$_.LowOrgName -eq $lowOrg} | Select -ExpandProperty "Number of Licenses Utilized"
                $counter++
            }
            else {
                Write-Host "$lowOrg not in sheet:" $sheet.MonthFile -ForegroundColor Yellow
            }
            
        }

        if ($counter -ne 0 ) {
            $avg = [int](($avg + $row.'Number of Licenses Utilized') / ($counter + 1))
        }

        $row."Monthly Change" = [int]$row."Number of Licenses Utilized" - [int]$row."Last Month"

        $row.'6 Month Average' = $avg
    }

    $newestSheet.File | Export-Csv -Path "./averages/$curMonth.csv" -NoTypeInformation

    return $newestSheet.File
}
