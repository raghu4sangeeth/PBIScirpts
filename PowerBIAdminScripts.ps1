# Note that downloading pbix would require workspace admin access and the below script adds the current user to admin role automatically. 

#####################################################
#           Helper Functions to Dump Error          #
#####################################################
function Get-ErrorInformation {
    [cmdletbinding()]
    param($incomingError)

    if ($incomingError -and (($incomingError | Get-Member | Select-Object -ExpandProperty TypeName -Unique) -eq 'System.Management.Automation.ErrorRecord')) {

        Write-Host `n"Error information:"`n
        Write-Host `t"Exception type for catch: [$($IncomingError.Exception | Get-Member | Select-Object -ExpandProperty TypeName -Unique)]"`n 

        if ($incomingError.InvocationInfo.Line) {
        
            Write-Host `t"Command                 : [$($incomingError.InvocationInfo.Line.Trim())]"
        
        }
        else {

            Write-Host `t"Unable to get command information! Multiple catch blocks can do this :("`n

        }

        Write-Host `t"Exception               : [$($incomingError.Exception.Message)]"`n
        Write-Host `t"Target Object           : [$($incomingError.TargetObject)]"`n

    }

    Else {

        Write-Host "Please include a valid error record when using this function!" -ForegroundColor Red -BackgroundColor DarkBlue

    }

}

#####################################################
#           Download PBIX reports                   #
#####################################################
function Download_PBIXReports {
    [cmdletbinding()]
    param($workspaceName)
    
    # #Connect to PBI Service Account
    # $userProfile = Connect-PowerBIServiceAccount –Environment Public 	
    
       
    if ([string]::IsNullOrWhitespace($workspaceName)) {
        #Retrieve all the workspaces as an admin
        $Workspaces = Get-PowerBIWorkspace -Scope Organization -All
        $tenantLevel = $true
    }
    else {
        #Retrieve the workspaces with the given workspace name
        $Workspaces = @( Get-PowerBIWorkspace  -Name $workspaceName -Scope Organization )
    }


    #Use seperate post fix to add to all the exports using current date 
    $CurrentDate = Get-Date –Format "yyyyMMdd" 

    # Build the outputpath.
    $folder = Read-Host -Prompt 'Input the folder name to save the reports(eg: C:\Out)'
    $OutPutPath = $folder + "\" + $CurrentDate 

    #Now loop through the workspaces, hence the ForEach
    ForEach ($Workspace in $Workspaces) {
            
        #only keep alphabets in the file / folder names
        $pattern = '\W'
        $folderName = $Workspace.name -replace $pattern
            
        #create the required rights to download 
        if ($tenantLevel) {
            # for now we are not checking if the user is already a member nor we are not removing the user once added. 
            Add-PowerBIWorkspaceUser -Scope Organization -Id $Workspace.Id -UserEmailAddress $userProfile.UserName -AccessRight Admin
        }
            
        #For all workspaces there is a new Folder destination: Outputpath + Workspacename
        $Folder = $OutPutPath + "\" + $folderName 
        #If the folder doens't exists, it will be created.
        If (!(Test-Path $Folder)) {
            New-Item –ItemType Directory –Force –Path $Folder
        }
        #At this point, there is a folder structure with a folder for all your workspaces 
	
	
        #Collect all (or one) of the reports from one or all workspaces 
        $PBIReports = Get-PowerBIReport –WorkspaceId $Workspace.Id -Scope Organization 						 

        #Now loop through these reports: 
        ForEach ($Report in $PBIReports) {
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
            Export-PowerBIReport –WorkspaceId $Workspace.ID –Id $Report.ID –OutFile $OutputFile
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
    #$userProfile = Connect-PowerBIServiceAccount –Environment Public 	
    
       
    if ([string]::IsNullOrWhitespace($workspaceName)) {
        #Retrieve all the workspaces as an admin
        $Workspaces = Get-PowerBIWorkspace -Scope Organization -All
    }
    else {
        #Retrieve the workspaces with the given workspace name
        $Workspaces = @( Get-PowerBIWorkspace  -Name $workspaceName -Scope Organization )
    }


    #Use seperate post fix to add to all the exports using current date 
    $CurrentDate = Get-Date –Format "yyyyMMdd" 

    # Build the outputpath.
    $folder = Read-Host -Prompt 'Input the folder name to save the Users(eg: C:\Out)'
    $OutPutPath = $folder + "\" + $CurrentDate
    $OutputFile = $OutPutPath + "\Users.csv" 

    #Now loop through the workspaces, hence the ForEach
    ForEach ($Workspace in $Workspaces) {
            
        #Important Note:
        # Commenting as we only need one users file. only keep alphabets in the file / folder names. 
        # when we hit throttling limits check if file exists before to move forward with the api call.
        #End.

        #$pattern = '\W'
        #$folderName = $Workspace.name -replace $pattern
            
        # #For all workspaces there is a new Folder destination: Outputpath + Workspacename
        # $Folder = $OutPutPath + "\" + $folderName 
        # #If the folder doens't exists, it will be created.
        # If (!(Test-Path $Folder)) {
        #     New-Item –ItemType Directory –Force –Path $Folder
        # }
        #At this point, there is a folder structure with a folder for all your workspaces 
	
        #Collect all (or one) of the reports from one or all workspaces 
        $PBIReports = Get-PowerBIReport –WorkspaceId $Workspace.Id -Scope Organization 						 

        #Now loop through these reports: 
        ForEach ($Report in $PBIReports) {
            #Important Note:
        # Commenting as we only need one users file. only keep alphabets in the file / folder names. 
        # when we hit throttling limits check if file exists before to move forward with the api call.
        #End.
            #$fileName = $Report.name -replace $pattern
            #The final collection including folder structure + file name is created.
            #$OutputFile = $OutPutPath + "\" + $folderName + "\" + $fileName + ".csv"
            
            #Your PowerShell comandline will say Downloading Workspacename Reportname
            Write-Host "Downloading "$Workspace.name":" $Report.name " Users."
            
            $url = "/admin/reports/" + $Report.ID + "/users"            
            $responseUsers = Invoke-PowerBIRestMethod -Url $url -Method Get 

            if ($responseUsers) {
                ForEach ($user in ($responseUsers | ConvertFrom-Json).value) {
                    $user | Select-Object @{Name='Workspace';Expression={$Workspace.name}}, @{Name='Report';Expression={$Report.name}}, * | Export-CSV $OutputFile -NoTypeInformation -Append -Force
                }
            }
           
        }
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
    else{
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
                Write-Host "Successully retrieved name: " $wsName

                #Move the workspace to shared capacity 
                Set-PowerBIWorkspace -Scope Organization -Id $ws.Id -CapacityId $capacityID

            }
        }
        elseif ($nameHeader -in $headers) {
            foreach ($ws in $wsInfo) {
                Write-Host "Working on the Workspace Name: " $ws.Name

                $wsId = (Get-PowerBIWorkspace -Scope Organization -Name $ws.Name).Id
                Write-Host "Successully retrieved Id: " $wsId

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
    $userProfile = Connect-PowerBIServiceAccount –Environment Public 	
    
    #add here the functions you would want to call
    #Add_WorkspacesToPremium
}
catch {
    Get-ErrorInformation -incomingError $_
    continue
}
Resolve-PowerBIError
