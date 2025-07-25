#Connect to PBI using service account
Connect-PowerBiServiceAccount
$accessToken=Get-PowerBIAccessToken -AsString
$headers-@{"Authorization"=$accessToken}

[void] [System.Reflection.Assembly]::LoadwithPartialName("System.Windows.Forms")
[void] [Reflection.Assembly]::LoadwithPartialName('Microsoft.VisualBasic')

Clear-Host
$title = 'New WorkspaceName'
$msg = 'Please provide a name you want to create a workspace: '
$WorkspaceName = [Microsoft.VisualBasic.Interaction]:: InputBox($msg, $title)

#Check for errors
If([string]:: IsNullOrEmpty($workspaceName )){
    [void] [System.Windows.Forms.MessageBox]:: show("Cannot create workspace with No Name", "Error Box")
    exit
}

$existingWorkspace = Get-PowerBIWorkspace -Name $workspaceName
if($existingWorkspace) {
    Write-Host "Workspace SworkspaceName exists" -BackgroundColor Yellow
    [void] [System.Windows.Forms.MessageBox]:: show("Workspace Already Exists...!!!!!", "Error")
    }
else {
    $Newworkspace=New-PowerBIworkspace -Name $workspaceName
    Write-Host "Workspace '$($Newworkspace.name)' created"
    [void] [System.Windows.Forms.MessageBox]:: show("Workspace Created Successfully", "Confirmation")
}
Disconnect-PowerBIServiceAccount
