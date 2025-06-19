# Note that downloading pbix would require workspace admin access and the below script adds the current user to admin role automatically. 

#####################################################
#           Helper Functions to Dump Error          #
#####################################################
function Get-ErrorInformation {
    [cmdletbinding()]
    param($incomingError)

    if ($incomingError -and (($incomingError | Get-Member | Select-Object -ExpandProperty TypeName -Unique) -eq 'System.Management.Automation.ErrorRecord')) {

        Write-Host `n"Error information:"`n
        Write-Host `t"Exception type for catch: [$($incomingError.Exception | Get-Member | Select-Object -ExpandProperty TypeName -Unique)]"`n 

        if ($incomingError.InvocationInfo.Line) {
        
            Write-Host `t"Command                 : [$($incomingError.InvocationInfo.Line.Trim())]"
        
        }
        else {

            Write-Host `t"Unable to get command information! Multiple catch blocks can do this :("`n

        }

        Write-Host `t"Exception               : [$($incomingError.Exception.Message)]"`n
        Write-Host `t"Target Object           : [$($incomingError.TargetObject)]"`n

    }

    else {

        Write-Host "Please include a valid error record when using this function!" -ForegroundColor Red -BackgroundColor DarkBlue

    }

}
#####################################################
#           Download AuditLogs                      #
#####################################################
function Download_AuditLogs {
    [cmdletbinding()]
    param($numberOfDays)
    
    # #Connect to PBI Service Account
    # $userProfile = Connect-PowerBIServiceAccount -Environment Public 	

    if ([string]::IsNullOrWhitespace($numberOfDays)) {
        return
    }

    # Build the outputpath.
    #Use seperate post fix to add to all the exports using current date 
    $CurrentDate = Get-Date -Format "yyyyMMdd" 
    $folder = Read-Host -Prompt 'Input the folder name to save the logs(eg: C:\Out)'
    $outPutPath = $folder + "\" + $CurrentDate 

    #If the folder doens't exists, it will be created.
    if (!(Test-Path $outPutPath)) {
        New-Item -ItemType Directory -Force -Path $outPutPath
    }
    
    # Number of days is 30 max for activity logs

    $numberOfDays..1 |
    ForEach-Object {
        $Date = (((Get-Date).Date).AddDays(-$_))
        $StartDate = (Get-Date -Date ($Date) -Format yyyy-MM-ddTHH:mm:ss)
        $EndDate = (Get-Date -Date ((($Date).AddDays(1)).AddMilliseconds(-1)) -Format yyyy-MM-ddTHH:mm:ss)
        
        Get-PowerBIActivityEvent -StartDateTime $StartDate -EndDateTime $EndDate -ResultType JsonString | 
        Out-File -FilePath "$OutPutPath\PowerBI_AuditLog_$(Get-Date -Date $Date -Format yyyyMMdd).json"
    }
}
#####################################################
#           Download PBIX reports                   #
#####################################################
function Grant-AdminAccessToWorkspace {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$WorkspaceId,

        [Parameter(Mandatory = $true)]
        [string]$ApiHostUri
    )

    # Get access token header
    $headers = Get-PowerBIAccessToken

    # Add required headers
    $headers["Referer"] = "https://app.powerbi.com"
    $headers["Origin"] = "https://app.powerbi.com"
    $headers["X-PowerBI-HostEnv"] = "Power BI Web App"
    $headers["X-PowerBI-User-Admin"] = "True"

    # Construct API URI
    $uri = "$ApiHostUri/metadata/admin/workspaces/$WorkspaceId/adminAccess"

    # Make PUT request
    Invoke-RestMethod -Uri $uri -Headers $headers -Method Put

    Write-Host "Admin access granted to workspace ID $WorkspaceId"
}

function Get-PowerBIReportWithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WorkspaceId,
        [int]$MaxRetries = 30,
        [int]$RetryDelaySeconds = 3600
    )
    $retryCount = 0
    while ($retryCount -lt $MaxRetries) {
        try {
            return Get-PowerBIReport -WorkspaceId $WorkspaceId -Scope Organization
        }
        catch {
            if ($_.Exception.Response.StatusCode.value__ -eq 429) {
                Write-Host "Received HTTP 429 (Too Many Requests). Sleeping for 1 hour before retrying..." -ForegroundColor Yellow
                Start-Sleep -Seconds $RetryDelaySeconds
                $retryCount++
            }
            else {
                throw $_
            }
        }
    }
    Write-Host "Failed to retrieve reports for workspace $WorkspaceId after $MaxRetries retries." -ForegroundColor Red
    return $null
}

function Export-PowerBIReportWithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WorkspaceId,
        [Parameter(Mandatory = $true)]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [string]$OutFile,
        [int]$MaxRetries = 30,
        [int]$RetryDelaySeconds = 3600
    )
    $retryCount = 0
    while ($retryCount -lt $MaxRetries) {
        try {
            Export-PowerBIReport -WorkspaceId $WorkspaceId -Id $ReportId -OutFile $OutFile
            return
        }
        catch {
            if ($_.Exception.Response.StatusCode.value__ -eq 429) {
                Write-Host "Received HTTP 429 (Too Many Requests) during export. Sleeping for 1 hour before retrying..." -ForegroundColor Yellow
                Start-Sleep -Seconds $RetryDelaySeconds
                $retryCount++
            }
            else {
                throw $_
            }
        }
    }
    Write-Host "Failed to export report $ReportId after $MaxRetries retries." -ForegroundColor Red
}

<#
+     .SYNOPSIS
+     Downloads PBIX reports from Power BI workspaces.
+ 
+     .PARAMETER workspaceName
+     Name of the workspace to download reports from. If not specified, all workspaces are processed.
+ 
+     .PARAMETER ApiHostUri
+     The Power BI API host URI for your region (e.g., https://wabi-west-us3-a-primary-redirect.analysis.windows.net/).
+     This is required for granting admin access to personal workspaces. Ensure you use the correct URI for your tenant's region!
+     #>
function Download_PBIXReports {
    [cmdletbinding()]
    param(
        [string] $WorkspaceName,
        [string] $ApiHostUri = "https://wabi-west-us3-a-primary-redirect.analysis.windows.net/")
    
    # Validate the userProfile
    if (-not $userProfile) {
        Write-Host "You are not connected to Power BI Service. Connecting now." -ForegroundColor Yellow
        $userProfile = Connect-PowerBIServiceAccount -Environment Public
    }
       
    if ([string]::IsNullOrWhitespace($WorkspaceName)) {
        #Retrieve all the workspaces as an admin
        $Workspaces = Get-PowerBIWorkspace -Scope Organization -All
        $tenantLevel = $true
    }
    else {
        #Retrieve the workspaces with the given workspace name
        $Workspaces = @( Get-PowerBIWorkspace -Name $WorkspaceName -Scope Organization )
    }


    #Use seperate post fix to add to all the exports using current date 
    $CurrentDate = Get-Date -Format "yyyyMMdd" 

    # Build the outputpath.
    $folder = Read-Host -Prompt 'Input the folder name to save the reports(eg: C:\Out)'
    $OutPutPath = $folder + "\" + $CurrentDate 

    #Now loop through the workspaces, hence the ForEach
    foreach ($Workspace in $Workspaces) {
        #only keep alphabets in the file / folder names
        $pattern = '\W'
        $folderName = $Workspace.name -replace $pattern
            
        #create the required rights to download 
        if ($tenantLevel) {
            if ($Workspace.Type -eq "PersonalGroup" -and $Workspace.Name -ne "My Workspace") {
                #Grant admin access to the personal workspace if the user is not already an admin - This is unsupported way!!
                #ensure you replace the ApiHostUri with the correct one for your region
                Grant-AdminAccessToWorkspace -WorkspaceId $Workspace.Id -ApiHostUri $ApiHostUri
            }
            else {
                #Write-Host "The workspace: " $Workspace.name " is read-only. Skipping admin access." -ForegroundColor Yellow
                # for now we are not checking if the user is already a member nor we are not removing the user once added. 
                Add-PowerBIWorkspaceUser -Scope Organization -Id $Workspace.Id -UserEmailAddress $userProfile.UserName -AccessRight Admin
            }
        }
            
        #For all workspaces there is a new Folder destination: Outputpath + Workspacename
        $Folder = $OutPutPath + "\" + $folderName 
        
        #If the folder doens't exists, it will be created.
        if (!(Test-Path $Folder)) {
            New-Item -ItemType Directory -Force -Path $Folder
        }
        
        #At this point, there is a folder structure with a folder for all your workspaces 
        $downloaded = $false
        #Collect all (or one) of the reports from one or all workspaces 
        $PBIReports = Get-PowerBIReportWithRetry -WorkspaceId $Workspace.Id

        #Now loop through these reports: 
        foreach ($Report in $PBIReports) {
            $fileName = $Report.name -replace $pattern

            #Your PowerShell comandline will say Downloading Workspacename Reportname
            Write-Host "Downloading "$Workspace.name":" $Report.name "to " $folderName":"$fileName
			
            #The final collection including folder structure + file name is created.
            $OutputFile = $OutPutPath + "\" + $folderName + "\" + $fileName + ".pbix"
			
            # If the file exists, delete it first; otherwise, the Export-PowerBIReport will fail.
            if (Test-Path $OutputFile) {
                Remove-Item $OutputFile
            }
			
            #The pbix is now really getting downloaded
            Export-PowerBIReportWithRetry -WorkspaceId $Workspace.Id -Id $Report.Id -OutFile $OutputFile

            if (Test-Path $OutputFile) {
                Write-Host "Successfully downloaded." -ForegroundColor Green
                $downloaded = $true
            }
            else {
                Write-Host "Failed to download the report: " $Report.name " from workspace: " $Workspace.name -ForegroundColor Red
            }
        }
        
        # If no reports were downloaded, rename the folder to indicate no reports were found
        if (!$downloaded) {
            Rename-Item -Path $Folder -NewName ("_" + $folderName)
        }
    }
}

#####################################################
#          Get all Users for Reports                #
#####################################################
function Get_ReportUsers {
    [cmdletbinding()]
    param($workspaceName)
    
    #Connect to PBI Service Account
    #$userProfile = Connect-PowerBIServiceAccount -Environment Public 	
    
       
    if ([string]::IsNullOrWhitespace($workspaceName)) {
        #Retrieve all the workspaces as an admin
        $Workspaces = Get-PowerBIWorkspace -Scope Organization -All
    }
    else {
        #Retrieve the workspaces with the given workspace name
        $Workspaces = @( Get-PowerBIWorkspace -Name $workspaceName -Scope Organization )
    }


    #Use seperate post fix to add to all the exports using current date 
    $CurrentDate = Get-Date -Format "yyyyMMdd" 

    # Build the outputpath.
    $folder = Read-Host -Prompt 'Input the folder name to save the Users(eg: C:\Out)'
    $OutPutPath = $folder + "\" + $CurrentDate
    $OutputFile = $OutPutPath + "\Users.csv" 

    # Initialize the array to collect all users
    $allUsers = @()

    #Now loop through the workspaces, hence the ForEach
    foreach ($Workspace in $Workspaces) {
        #Important Note:
        # Commenting as we only need one users file. only keep alphabets in the file / folder names. 
        # when we hit throttling limits check if file exists before to move forward with the api call.
        #End.

        #Collect all (or one) of the reports from one or all workspaces 
        $PBIReports = Get-PowerBIReport -WorkspaceId $Workspace.Id -Scope Organization 						 

        #Now loop through these reports: 
        foreach ($Report in $PBIReports) {
            
            #Your PowerShell comandline will say Downloading Workspacename Reportname
            Write-Host "Downloading "$Workspace.name":" $Report.name " Users."
            
            $url = "/admin/reports/" + $Report.ID + "/users"            
            $responseUsers = Invoke-PowerBIRestMethod -Url $url -Method Get 

            if ($responseUsers) {
                foreach ($user in ($responseUsers | ConvertFrom-Json).value) {
                    $allUsers += $user | Select-Object @{Name = 'Workspace'; Expression = { $Workspace.name } }, @{Name = 'Report'; Expression = { $Report.name } }, *
                }
            }
        }
    }
    if ($allUsers.Count -gt 0) {
        $allUsers | Export-Csv $OutputFile -NoTypeInformation -Force
    }
}

#####################################################
#           Assign workspaces to Premium            #
#####################################################
function Add_WorkspacesToPremium {
    [cmdletbinding()]
    param($filePath, $capacityName)

    if ([string]::IsNullOrWhitespace($filePath)) {
        #CSV File Format - Single Column CSV with header being either Id or Name of workspace. 
        $filePath = Read-Host -Prompt 'Input the file path of workspace ids saved as CSV (eg: C:\Out\Workspaces.csv)'
    }

    if ([string]::IsNullOrWhitespace($capacityName)) {
        $capacityName = Read-Host -Prompt 'Input the capacityName'
    }

    $capacityID = (Get-PowerBICapacity -Scope Organization | Where-Object { $_.DisplayName -eq $capacityName }).Id
    if ([string]::IsNullOrWhitespace($capacityID)) {
        Write-Host "No capacity with the name: " $capacityName " exists. Defaulting to Shared."
        $capacityID = "00000000-0000-0000-0000-000000000000"
    }
    else {
        Write-Host "Found the capacity: " $capacityName " (ID: " $capacityID ")"
    }

    if (Test-Path $filePath) {

        #Identify the workspace to be moved to premium capacity
        $wsInfo = Import-Csv -Path $filePath 
        $headers = ($wsInfo | Get-Member -MemberType NoteProperty).Name
        $idHeader = "Id"
        $nameHeader = "Name"

        if ( $idHeader -in $headers ) {
            foreach ($ws in $wsInfo) {
                Write-Host "Working on the Workspace ID: " $ws.Id

                $wsName = (Get-PowerBIWorkspace -Scope Organization -Id $ws.Id).Name
                Write-Host "Successfully retrieved name: " $wsName

                #Move the workspace to shared capacity 
                Set-PowerBIWorkspace -Scope Organization -Id $ws.Id -CapacityId $capacityID

            }
        }
        elseif ($nameHeader -in $headers) {
            foreach ($ws in $wsInfo) {
                Write-Host "Working on the Workspace Name: " $ws.Name

                $wsId = (Get-PowerBIWorkspace -Scope Organization -Name $ws.Name).Id
                Write-Host "Successfully retrieved Id: " $wsId

                #Move the workspace to shared capacity 
                Set-PowerBIWorkspace -Scope Organization -Id $wsId -CapacityId $capacityID
            }
        }

    }
    else {
        Write-Host "Invalid File Path."
    }
}


#####################################################
#                   Main Calls                      #
##################################################### 
try {
    #Connect to PBI Service Account
    $userProfile = Connect-PowerBIServiceAccount -Environment Public
    #add here the functions you would want to call
    Download_PBIXReports
}
catch {
    Get-ErrorInformation -incomingError $_
    return
} 
#Resolve-PowerBIError
finally {
    #Disconnect from Power BI Service Account
    if ($userProfile) {
        Disconnect-PowerBIServiceAccount -Force
    }
}