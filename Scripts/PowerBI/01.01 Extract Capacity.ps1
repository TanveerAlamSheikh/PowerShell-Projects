#param(
#    [String]$Email_ID,
#    [String]$User_Pwd,
#    [String]$Folder_Path,
#    [String]$Requestor
#)
CLS

$accessToken = ""
Connect-PowerBiServiceAccount
$accessToken=Get-PowerBIAccessToken -AsString
$headers-@{"Authorization"=$accessToken}
Clear-Host
<#
#**********************************Change the User Name and Password*******************************
    $PWord = ConvertTo-SecureString -String $User_Pwd -AsPlainText -Force
    $UserCred = New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $Email_ID, $PWord
    #$UserCred = New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $Email_ID, $User_Pwd
    If(Connect-PowerBIServiceAccount -Credential $UserCred -ErrorAction SilentlyContinue) {
        $accessToken=Get-PowerBIAccessToken -AsString
        $headers-@{"Authorization"=$accessToken}
    }Else{
        Write-Host "Connection Failed - Invalid Credentials" -BackgroundColor Red
        [Void] [Microsoft.VisualBasic.Interaction]:: MsgBox("Invalid Credentails", 'OKOnly, SystemModal, Critical', 'Connection Failed')
        Exit
    }
#>

$today = (Get-Date).ToString('yyyyMMdd_HHmmss')
if(Test-Path $Folder_Path\CapacityLog.txt) {Remove-Item $Folder_Path\CapacityLog.txt -Verbose}
Write-Output "Status`tScript_Run_By`tError_Message" | Out-File $Folder_Path\CapacityLog.txt -Append

Cls
Write-Host "***********************Extracting Capacity***************************" -BackgroundColor DarkGray -ForegroundColor Yellow
Try{
    $apiURLAdd="https://api.powerbi.com/v1.0/myorg/admin/capacities?$expand=tenantKey"
    #$apiURLAdd-"https://api.powerbi.com/v1.0/myorg/admin/capacities"
    Start-Sleep -Seconds 10
    $response=Invoke-RestMethod -Uri $apiURLAdd -Headers $headers -Method Get
    $ResponseValues = $response.value | Select-Object id, displayName, sku, state, region, admins, capacityUserAccessRight, tenanKeyId, tenantKey
    Write-Host ("Exported") -BackgroundColor DarkGreen -ForegroundColor White
    $ResponseValues | Export-Csv -path "$Folder_Path\Capacity_$today.csv" -NoTypeInformation
    Write-Output "Success`t$Requestor" | Out-File -FilePath $Folder_Path\CapacityLog.txt -Append
}
Catch{
    Write-Host ("Failed") + $_.exception.Message -BackgroundColor Red
    Write-Output "Failed`t$Requestor`t$_.exception.Message" | Out-File -FilePath $Folder_Path\CapacityLog.txt -Append
}

Disconnect-PowerBIServiceAccount
[Void] [System.Windows.Forms.MessageBox]:: show("Script Run Completed..!!","Done")
Start-Process 'C:\WINDOWS\system32\notepad.exe'$Folder_Path\CapacityLog.txt -WindowStyle Maximized


