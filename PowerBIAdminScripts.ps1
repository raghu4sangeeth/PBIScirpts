# Note that downloading pbix would require workspace admin access and the below script adds the current user to admin role automatically. 

#####################################################
#           Helper Functions to Dump Error          #
#####################################################
# Set the log file path
$LogFile = Join-Path -Path (Split-Path -Parent $MyInvocation.MyCommand.Path) -ChildPath "PowerBIAdminScripts.log"

# Declare userProfile as a script-scoped variable to avoid scope issues
$script:userProfile = $null

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [string]$Color = "White"
    )
    # Write to console
    Write-Host $Message -ForegroundColor $Color
    # Write to log file (without color)
    Add-Content -Path $LogFile -Value ("[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message")
}

function Get-ErrorInformation {
    [cmdletbinding()]
    param($incomingError)

    if ($incomingError -and (($incomingError | Get-Member | Select-Object -ExpandProperty TypeName -Unique) -eq 'System.Management.Automation.ErrorRecord')) {

        Write-Log "Error information:"
        Write-Log "Exception type for catch: [$($incomingError.Exception | Get-Member | Select-Object -ExpandProperty TypeName -Unique)]"

        if ($incomingError.InvocationInfo.Line) {
            Write-Log "Command                 : [$($incomingError.InvocationInfo.Line.Trim())]"
        }
        else {
            Write-Log "Unable to get command information! Multiple catch blocks can do this :("
        }

        Write-Log "Exception               : [$($incomingError.Exception.Message)]"
        Write-Log "Target Object           : [$($incomingError.TargetObject)]"
    }
    else {
        Write-Log "Please include a valid error record when using this function!" "Red"
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
    try {
        Invoke-RestMethod -Uri $uri -Headers $headers -Method Put
        Write-Log "[INFO] Admin access granted to workspace ID $WorkspaceId" "Green"
    }
    catch {
        Write-Log "[WARN] Admin access cant be granted to workspace ID $WorkspaceId. (mostly because you already have access)" "Yellow"
    }
}

function Get-PowerBIReportWithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WorkspaceId,
        [int]$MaxRetries = 30,
        [int]$RetryDelaySeconds = 05
    )
    $retryCount = 0
    # Start with a delay
    Start-Sleep -Seconds $RetryDelaySeconds

    while ($retryCount -lt $MaxRetries) {
        try {
            return Get-PowerBIReport -WorkspaceId $WorkspaceId -Scope Organization
        }
        catch {
            if ($_.Exception.Response.StatusCode.value__ -eq 429) {
                Write-Log "[WARN] Received HTTP 429 (Too Many Requests). Sleeping for 1 hour before retrying..." "Yellow"
                Start-Sleep -Seconds ($RetryDelaySeconds * 60)
                $retryCount++
            }
            else {
                throw $_
            }
        }
    }
    Write-Log "[ERRO] Failed to retrieve reports for workspace $WorkspaceId after $MaxRetries retries." "Red"
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
        [int]$RetryDelaySeconds = 05
    )
    $retryCount = 0
    # Start with a delay
    Start-Sleep -Seconds $RetryDelaySeconds

    while ($retryCount -lt $MaxRetries) {
        try {
            Export-PowerBIReport -WorkspaceId $WorkspaceId -Id $Id -OutFile $OutFile
            return
        }
        catch {
            if ($_.Exception.Response.StatusCode.value__ -eq 429) {
                Write-Log "[WARN] Received HTTP 429 (Too Many Requests) during export. Sleeping for 1 hour before retrying..." "Yellow"
                Start-Sleep -Seconds ($RetryDelaySeconds * 60)
                $retryCount++
            }
            else {
                Write-Log "[ERRO] Failed to export report: $OutFile." "Red"
                throw $_
            }
        }
    }
    Write-Log "[ERRO] Failed to export report $Id after $MaxRetries retries: $OutFile." "Red"
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
        Write-Log "[WARN] You are not connected to Power BI Service. Connecting now." "Yellow"
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
            Write-Log "[INFO] Downloading $($Workspace.name): $($Report.name) to ${folderName}: ${fileName}"
			
            #The final collection including folder structure + file name is created.
            $OutputFile = $OutPutPath + "\" + $folderName + "\" + $fileName + ".pbix"
			
            # If the file exists, delete it first; otherwise, the Export-PowerBIReport will fail.
            if (Test-Path $OutputFile) {
                Remove-Item $OutputFile
            }
			
            #The pbix is now really getting downloaded
            Export-PowerBIReportWithRetry -WorkspaceId $Workspace.Id -Id $Report.Id -OutFile $OutputFile

            if (Test-Path $OutputFile) {
                Write-Log "[INFO] Successfully downloaded." "Green"
                $downloaded = $true
            }
            else {
                Write-Log "[ERRO] Failed to download the report: $($Report.name) from workspace: $($Workspace.name)" "Red"
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
    #If the folder doens't exists, it will be created.
    if (!(Test-Path $OutPutPath)) {
        New-Item -ItemType Directory -Force -Path $OutPutPath
    }
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
            Write-Log "[INFO] Downloading $($Workspace.name): $($Report.name) Users."
            
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
        Write-Log "[WARN] No capacity with the name: $capacityName exists. Defaulting to Shared." "Yellow"
        $capacityID = "00000000-0000-0000-0000-000000000000"
    }
    else {
        Write-Log "[INFO] Found the capacity: $capacityName (ID: $capacityID)" "Green"
    }

    if (Test-Path $filePath) {

        #Identify the workspace to be moved to premium capacity
        $wsInfo = Import-Csv -Path $filePath 
        $headers = ($wsInfo | Get-Member -MemberType NoteProperty).Name
        $idHeader = "Id"
        $nameHeader = "Name"

        if ( $idHeader -in $headers ) {
            foreach ($ws in $wsInfo) {
                Write-Log "[INFO] Working on the Workspace ID: $($ws.Id)" "Green"

                $wsName = (Get-PowerBIWorkspace -Scope Organization -Id $ws.Id).Name
                Write-Log "[INFO] Successfully retrieved name: $wsName" "Green"

                #Move the workspace to shared capacity 
                Set-PowerBIWorkspace -Scope Organization -Id $ws.Id -CapacityId $capacityID

            }
        }
        elseif ($nameHeader -in $headers) {
            foreach ($ws in $wsInfo) {
                Write-Log "[INFO] Working on the Workspace Name: $($ws.Name)" "Green"

                $wsId = (Get-PowerBIWorkspace -Scope Organization -Name $ws.Name).Id
                Write-Log "[INFO] Successfully retrieved Id: $wsId" "Green"

                #Move the workspace to shared capacity 
                Set-PowerBIWorkspace -Scope Organization -Id $wsId -CapacityId $capacityID
            }
        }

    }
    else {
        Write-Log "[ERRO] Invalid File Path." "Red"
    }
}


#####################################################
#                   Main Calls                      #
##################################################### 
try {
    #Connect to PBI Service Account
    $userProfile = Connect-PowerBIServiceAccount -Environment Public
    #add here the functions you would want to call
    Get_ReportUsers
}
catch {
    Get-ErrorInformation -incomingError $_
    return
} 
#Resolve-PowerBIError
finally {
    #Disconnect from Power BI Service Account
    if ($userProfile) {
        Disconnect-PowerBIServiceAccount
    }
}