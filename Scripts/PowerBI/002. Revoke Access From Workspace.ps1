#Connect to PBI using service account
Connect-PowerBiServiceAccount
$accessToken=Get-PowerBIAccessToken -AsString
$headers=@{"Authorization"=$accessToken}

#[void] [System.Reflection.Assembly]:: LoadWithPartialName("System.Drawing")
#[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
#[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
Add-Type-AssemblyName System.Windows.Forms
Add-Type-AssemblyName System.Drawing
Add-Type-AssemblyName Microsoft.VisualBasic

function MyCustomForm ([string]$title, [string]$Label1, [string]$Label2)
{
#[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
#[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
    $objForm = New-Object System.Windows.Forms.Form
    $objForm.Text = $title
    $objForm.Size = New-Object System.Drawing.Size(400,200)
    $objForm.StartPosition = "CenterScreen"
    $objForm.KeyPreview = $True
    $objForm.Add_KeyDown({
            if ($_.KeyCode -eq "Enter" -or $_.KeyCode -eq "Escape")
            {
                $objForm.Close()
            }
        })

    $OKButton = New-Object System.Windows.Forms.Button
    #$CancelButton = New-Object System.Windows.Forms.Button
    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel2 = New-Object System.Windows.Forms.Label
    $objTextBox = New-Object System.Windows.Forms.TextBox
    $objTextBox2 = New-Object System.Windows.Forms.TextBox

    $OKButton.Location = New-Object System.Drawing.Size(75,130)
    #$CancelButton.Location = New-Object System.Drawing.Size(190,130)
    $objLabel.Location = New-Object System.Drawing.Size(10,20)
    $objLabel2.Location = New-Object System.Drawing.Size(10,70)
    $objTextBox.Location = New-Object System.Drawing.Size(10,40)
    $objTextBox2.Location = New-Object System.Drawing.Size(10,90)

    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    #$CancelButton.Size = New-Object System.Drawing.Size(75,23) 
    $objLabel.Size = New-Object System.Drawing.Size (280,20)
    $objLabel2.Size = New-Object System.Drawing.Size(280,20)
    $objTextBox.Size = New-Object System.Drawing.Size(350,20)
    $objTextBox2.Size = New-Object System.Drawing.Size(350,20)

    $OKButton.Text = "Submit"
    #$CancelButton.Text = "Cancel"
    $objLabel.Text = $Label1 + "*"
    $objLabel2.Text = $Label2 + "*"

    $OKButton.Add_Click({$objForm.Close()})
    #$CancelButton.Add_Click({$objForm.Close()})

    $objForm.Controls.Add($OKButton)
    #$objForm.Controls.Add($CancelButton) 
    $objForm.Controls.Add($objLabel)
    $objForm.Controls.Add($objLabel2)
    $objForm.Controls.Add($objTextBox)
    $objForm.Controls.Add($objTextBox2)

    $objForm.Topmost = $True
    $objForm.Add_Shown({$objForm.Activate()}) 
    $objForm.Add_Shown({$objTextBox.Select()})
    [void]$objForm.ShowDialog()


return $objTextBox.Text, $objTextBox2.Text
}
Clear-Host
$InputBox = [System.Windows.Forms.MessageBox]:: Show("Do you want to continue?", "Confirmation Required", [System.Windows.Forms.MessageBoxButtons]::YesNo)

switch ($InputBox) {
    No{
        Write-Host "No Action Taken......!" -ForegroundColor Red -BackgroundColor Yellow
        Exit
        }
    }

$return= MyCustomForm "Remove Workpace Access for the Users" "Workspace ID" "User Email ID"
$WorkspaceID = $return[0]

$EmailID= $return[1].ToLower()

#-----------------------------------------------------------------Check for errors
If([string]::IsNullOrEmpty($WorkspaceID) -and [string]:: IsNullOrEmpty($EmailID))
{
    Exit
}
    elseIf([string]:: IsNullOrEmpty($WorkspaceID))
    {
        [Microsoft.VisualBasic.Interaction]:: MsgBox("Please provide the Workspace ID.", 'OKOnly, SystemModal, Critical', 'Input not provided')
        Exit
    }
        elseIf([string]:: IsNullOrEmpty($EmailID))
        {
            [Microsoft.VisualBasic.Interaction]:: MsgBox("Please provide Email ID.", 'OKOnly, SystemModal, Critical', 'Input not provided')
            Exit
        }
            Elseif ($WorkspaceID.Length -gt 36 -or $WorkspaceID.Length -lt 36)
            {
                [Microsoft.VisualBasic.Interaction]:: MsgBox("Please check Workspace ID and provide the correct detail.`n`nGuid should contain 32 digits with 4 dashes. n(xxxXXXXXX-XXXX-XXXX-XXXX-XXXXXxxxxxxx)", 'OKOnly, SystemModal, Critical', 'Input not provided')
                Exit
            }
                $WS=Get-PowerBIWorkspace - Id $WorkspaceID -Scope organization 
                elseIf ($WS.Name.Length -eq 0)
                {
                    [Microsoft.VisualBasic.Interaction]::MsgBox("Provided Workspace ID doesn't exist in the tenant.", 'OKOnly, SystemModal, Critical', 'Input not provided')
                    Exit

                }
                    if ($EmailID.IndexOf("corebridgefinancial.com") -eq -1)
                    { 
                        [Microsoft.VisualBasic.Interaction]:: MsgBox("Access can only be provided to Corebridge User.nPlease provide the correct Email ID.", 'OKOnly, SystemModal, Critical', 'Input not provided')
                        Exit
                    }

Try{
    $apiURLAdd="https://api.powerbi.com/v1.0/myorg/groups/" + $WorkspaceID + "/users/" + $EmailID #"https://api.powerbi.com/v1.0/myorg/admin/groups/ $WorkspaceId/users"
    $X=Invoke-RestMethod -Uri $apiURLAdd -Body $body -Headers $headers -Method DELETE
    Write-Host "Access Revoked Successfully" -BackgroundColor Green
    #[void][System.Windows.Forms.MessageBox]:: show("Workspace access provided to :" + $Capacity, "Successful")
    [Microsoft.VisualBasic.Interaction]:: MsgBox("Workspace access revoked for :`n" + $EmailID, 'OKOnly, SystemModal, Information', 'Successful')
}
Catch{
    Write-Host "Access Couldn't be Revoked" -BackgroundColor Red
    #[void][System.Windows.Forms.MessageBox]:: show("Couldn't revoke the access from Workspace", "Failed") 
    [Microsoft.VisualBasic.Interaction]:: MsgBox("Couldn't revoke Workspace access forn" + $EmailID, 'OKOnly, SystemModal, Critical', 'Failed')
}
Disconnect-PowerBIServiceAccount