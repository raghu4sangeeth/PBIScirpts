$body       = '{capacityId:"xxx-xxxx-xxx-xxx"}'
$workspaces = Import-Csv -Path D:\WorkSpaces.csv -Delimiter ","
#Log in to service
Connect-PowerBIServiceAccount

$workspaces | ForEach-Object {
    
    $WorkspaceID = $_.Id
    $url = "groups/$WorkspaceID/AssignToCapacity"

    Write-Host $url
    Write-Host $body

    $re = Invoke-PowerBIRestMethod -Url $Url -Method Post -Body $body
    $re
}