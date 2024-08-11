<# :
    @echo off & chcp 65001 >nul & cd /d "%~dp0" & Title MSI Properties Viewer

    ::========= SETTINGS =========
    set "Powershell_WindowStyle=Hidden"  :: Normal, Hidden, Minimized, Maximized
    set "Show_Loading=true"              :: Show cmd while preparing powershell
    set "Ensure_Local_Running=true"      :: If not launched from disk 'C', Re-Write in %temp% then execute
        set "Show_Writing_Lines=false"   :: Show lines writing in %temp% while preparing powershell
        set "Debug_Writting_Lines=false" :: Pause between each line writing (press a key to see next line)
    ::============================
 
    if "%Show_Writing_Lines%"=="true" set "Show_Loading=true"
    if "%Debug_Writting_Lines%"=="true" set "Show_Loading=true" && set "Show_Writing_Lines=true"
    if "%Show_Loading%"=="false" (
        if not DEFINED IS_MINIMIZED set IS_MINIMIZED=1 && start "" /min "%~dpnx0" %* && exit
        ) else (if "%Show_Writing_Lines%"=="false" if "%Powershell_WindowStyle%"=="Hidden" mode con: cols=55 lines=3)
    echo. & echo  Loading...
;   if "%Ensure_Local_Running%"=="true" if "%~d0" NEQ "C:" ((
;       for /f "eol=; usebackq delims=" %%k in ("%~f0") do (
;           setlocal enabledelayedexpansion & set "line=%%k" & echo(!line!
;           if "%Show_Writing_Lines%"=="true" echo(!line! 1>&2
;           if "%Debug_Writting_Lines%"=="true" pause 1>&2 >nul
;           endlocal
;       )) > "%temp%\%~nx0" & start "" cmd.exe /c "%temp%\%~nx0" %* & exit)

    cls & echo. & echo  Launching PowerShell...
    powershell /nologo /noprofile /executionpolicy bypass /windowstyle %Powershell_WindowStyle% /command ^
        "&{[ScriptBlock]::Create((gc """%~f0""" -Raw)).Invoke(@(&{$args}%*))}"

    if "%~dp0" NEQ "%temp%\" (exit) else ((goto) 2>nul & del "%~f0")
#>



Add-Type -AssemblyName System.Windows.Forms  
Add-Type -AssemblyName System.Drawing  

$loadingForm = New-Object System.Windows.Forms.Form; $loadingForm.Text = "Loading..."
$loadingForm.Size = New-Object System.Drawing.Size(300,100); $loadingForm.StartPosition = "CenterScreen"
$loadingForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog; $loadingForm.ControlBox = $false
$launch_progressBar = New-Object System.Windows.Forms.ProgressBar; $launch_progressBar.Location = New-Object System.Drawing.Point(10,30)
$launch_progressBar.Size = New-Object System.Drawing.Size(260,20); $launch_progressBar.Style = "Continuous"
$loadingLabel = New-Object System.Windows.Forms.Label; $loadingLabel.Text = "Loading..."
$loadingLabel.Location = New-Object System.Drawing.Point(10,10); $loadingLabel.Size = New-Object System.Drawing.Size(280,20)
$loadingForm.Controls.Add($launch_progressBar); $loadingForm.Controls.Add($loadingLabel)
$loadingForm.Show(); $loadingForm.Refresh()

$launch_progressBar.Value = 10
$loadingLabel.Text = "Loading interface..."


function Get-MsiProperty {  
    [CmdletBinding()] 
    param (  
        [Parameter(Mandatory=$true)] 
        [ValidateScript({Test-Path $_})] 
        [string]$Path,  
  
        [Parameter(Mandatory=$true)] 
        [string]$Property  
    )  
  
    $WindowsInstaller = $null 
    $MSIDatabase = $null 
    $View = $null 
    $Record = $null 
  
    try { 
        $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer  
  
        # Ouverture de la base de données MSI en lecture seule 
        $MSIDatabase = $WindowsInstaller.OpenDatabase($Path, 0) 
  
        # Construction et exécution de la requête SQL 
        $Query = "SELECT Value FROM Property WHERE Property = '$([System.Security.SecurityElement]::Escape($Property))'" 
        $View = $MSIDatabase.OpenView($Query) 
        $View.Execute() 
  
        # Récupération du résultat de la requête 
        $Record = $View.Fetch() 
  
        if ($null -ne $Record) {  
            # Extraction de la valeur de la propriété 
            $Record.StringData(1) 
        } 
    }  
    catch { 
        Write-Error "An error occurred while retrieving the MSI property: $_" 
    } 
    finally { 
        # Fermeture et libération des objets COM 
        if ($null -ne $Record) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Record) | Out-Null } 
        if ($null -ne $View) {  
            $View.Close() 
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($View) | Out-Null  
        } 
        if ($null -ne $MSIDatabase) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($MSIDatabase) | Out-Null } 
        if ($null -ne $WindowsInstaller) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-Null } 
  
        # Forcer le ramasse-miettes pour libérer les ressources COM 
        [System.GC]::Collect() 
        [System.GC]::WaitForPendingFinalizers() 
    } 
} 

$launch_progressBar.Value = 20
$loadingLabel.Text = "Loading interface..."
  
$form = New-Object System.Windows.Forms.Form  
$form.Text = "MSI Properties extractor"  
$form.Size = New-Object System.Drawing.Size(600, 375)  
  
$labels = @()  
$textBoxes = @()  
$copyButtons = @()  
  
$properties = [ordered]@{  
    "GUID" = "ProductCode"  
    "Product Name" = "ProductName"  
    "Product Version" = "ProductVersion"  
    "Manufacturer" = "Manufacturer"  
    "Upgrade Code" = "UpgradeCode"  
}  
   
$yPos = 35  
foreach ($key in $properties.Keys) {  
    $label = New-Object System.Windows.Forms.Label  
    $label.Text = "$key :"  
    $label.AutoSize = $true  
    $label.Location = New-Object System.Drawing.Point(10, $yPos)  
    $form.Controls.Add($label)  
    $labels += $label  
  
    $textBox = New-Object System.Windows.Forms.TextBox  
    $textBox.Size = New-Object System.Drawing.Size(50, 25)  
    $textBox.Location = New-Object System.Drawing.Point(100, $yPos)  
    $textBox.ReadOnly = $true  
    $form.Controls.Add($textBox)  
    $textBoxes += $textBox  
  
    $copyButton = New-Object System.Windows.Forms.Button  
    $copyButton.Text = "COPY"  
    $copyButton.Size = New-Object System.Drawing.Size(60, 25)  
    $copyButton.Enabled = $false  
    $form.Controls.Add($copyButton)  
    $copyButtons += $copyButton  
  
    $copyButton.Add_Click({  
        $buttonIndex = $copyButtons.IndexOf($this)  
        [System.Windows.Forms.Clipboard]::SetText($textBoxes[$buttonIndex].Text)  
    })  
  
    $yPos += 30  
}  

$launch_progressBar.Value = 30
$loadingLabel.Text = "Loading interface..."

$separatorLine = New-Object System.Windows.Forms.Panel  
$separatorLine.Height = 1  
$separatorLine.Width = [int]$form.ClientSize.Width  
$separatorLine.BorderStyle = [System.Windows.Forms.BorderStyle]::None  
$separatorLine.BackColor = [System.Drawing.Color]::Gray  
$separatorLine.AutoSize = $true  
$separatorLine.Location = New-Object System.Drawing.Point(0, ($yPos+30))  
$form.Controls.Add($separatorLine)  
  
# Ajout du label pour MSI Path  
$labelMsiPath = New-Object System.Windows.Forms.Label  
$labelMsiPath.Text = "MSI Path:"  
$labelMsiPath.AutoSize = $true  
$labelMsiPath.Location = New-Object System.Drawing.Point(10, $yPos)  
$form.Controls.Add($labelMsiPath)  


  
$yPos += 25  # Augmenter la position Y pour le TextBox du chemin  
  
$textBoxPath = New-Object System.Windows.Forms.TextBox  
$textBoxPath.Size = New-Object System.Drawing.Size(260, 25)  
#$textBoxPath.Location = New-Object System.Drawing.Point(10, $yPos)  
$form.Controls.Add($textBoxPath)  
  
$progressBar = New-Object System.Windows.Forms.ProgressBar  
$progressBar.Size = New-Object System.Drawing.Size(260, 23)  
$progressBar.Style = "Continuous"  
$progressBar.Visible = $false  
$form.Controls.Add($progressBar)  
  
$findGuidButton = New-Object System.Windows.Forms.Button  
$findGuidButton.Text = "FIND GUID"  
$findGuidButton.Size = New-Object System.Drawing.Size(85, 25)  
$form.Controls.Add($findGuidButton)  
  
$browseButton = New-Object System.Windows.Forms.Button  
$browseButton.Text = "BROWSE"  
$browseButton.Size = New-Object System.Drawing.Size(85, 25)  
$form.Controls.Add($browseButton)  
  
$script:fromBrowseButton = $false  

$launch_progressBar.Value = 50
$loadingLabel.Text = "Loading interface..."

# Fonction pour ajuster la position des boutons  
function Update-ButtonPositions {  
    $formWidth = [int]$form.ClientSize.Width  
    $formHeight = [int]$form.ClientSize.Height  
  
    # Ajuster la largeur des TextBox  
    $textBoxWidth = $formWidth - 177  # 70 pixels de marge gauche + 40 pixels de marge droite  
    foreach ($textBox in $textBoxes) {  
        $textBox.Width = $textBoxWidth  
    }  
    $separatorLine.Width = [int]$form.ClientSize.Width  
  
    $buttonSpacing = [int](($formWidth - 2 * $findGuidButton.Width) / 3)  
    $browseButtonX = $buttonSpacing  
    $findGuidButtonX = 2 * $buttonSpacing + $browseButton.Width  
    $buttonY = $formHeight - 40  
  
    $browseButton.Location = New-Object System.Drawing.Point($browseButtonX, $buttonY)  
    $findGuidButton.Location = New-Object System.Drawing.Point($findGuidButtonX, $buttonY)  
  
    $textBoxPath.Location = New-Object System.Drawing.Point(10, ($buttonY - 30))  
    $textBoxPath.Width = $formWidth - 20  
    $progressBar.Width = $formWidth - 77  
  
    # Ajuster la position des boutons de copie  
    for ($i = 0; $i -lt $copyButtons.Count; $i++) {  
        $copyButtonX = $formWidth - 71  
        $copyButtonY = $textBoxes[$i].Location.Y - 4  
        $copyButtons[$i].Location = New-Object System.Drawing.Point($copyButtonX, $copyButtonY)  
    }  
  
    # Ajuster la position du label MSI Path  
    $labelMsiPathY = ($textBoxPath.Location.Y - 25)  
    $labelMsiPath.Location = New-Object System.Drawing.Point(10, $labelMsiPathY)  
    $progressBar.Location = New-Object System.Drawing.Point(67, ($labelMsiPathY - 3))  
}  
  
# Appeler la fonction initialement  
Update-ButtonPositions  
   
$form.Add_Resize({ Update-ButtonPositions })  

$launch_progressBar.Value = 60
$loadingLabel.Text = "Loading interface..."


# Ajouter un gestionnaire pour le bouton Parcourir  
$browseButton.Add_Click({  
  
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog  
    $openFileDialog.Filter = "MSI files (*.msi)|*.msi|All files (*.*)|*.*"  
    $openFileDialog.Title = "Select a MSI file"  
  
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {  
        $selectedFile = $openFileDialog.FileName  
        if ($selectedFile -match "\.msi$") {  
            if ($selectedFile -ne $textBoxPath.Text) {  
                foreach ($textBox in $textBoxes) {  
                    $textBox.Text = ""  
                }  
                foreach ($copyButton in $copyButtons) {  
                    $copyButton.Enabled = $false  
                }  
                $progressBar.Visible = $true  
                $progressBar.Value = 5  
                $job = Start-Job -ScriptBlock {  
                    param($path)  
                    Start-Sleep -Milliseconds 100  
                    return $path  
                } -ArgumentList $selectedFile  
  
                while ($job.State -eq 'Running') {  
                    if ($progressBar.Value -lt 90) {  
                        $progressBar.Value += 3  
                    }  
                    Start-Sleep -Milliseconds 50  
                }  
  
                $result = Receive-Job -Job $job  
                Remove-Job -Job $job  
  
                $script:fromBrowseButton = $true  
                $textBoxPath.Text = $result  
                $progressBar.Value = 100  
                $progressBar.Visible = $false  
  
                $findGuidButton.PerformClick()  
            } else {  
                [System.Windows.Forms.MessageBox]::Show("Preaching to the choir", "Same File", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)  
            }  
        }  
    }  
})  

$launch_progressBar.Value = 70
$loadingLabel.Text = "Loading interface..."

# Ajouter un gestionnaire pour le bouton Find GUID  
$findGuidButton.Add_Click({  
    $progressBar.Visible = $true  
    $progressBar.Value = 0  
    $msiPath = $textBoxPath.Text  
    $msiPath = $textBoxPath.Text.Trim('"')  
  
    $i = 0  
    $totalKeys = $properties.Keys.Count 
    $incrementValue = 100 / $totalKeys 
  
    foreach ($key in $properties.Keys) {  
        $progressBar.Value = [Math]::Round($progressBar.Value + $incrementValue) 
        $property = $properties[$key]  
  
        $value = Get-MsiProperty -Path $msiPath -Property $property  
        if ($value) {  
            $textBoxes[$i].Text = $value  
            $textBoxes[$i].Text = $textBoxes[$i].Text.Trim()  
            $copyButtons[$i].Enabled = $true  
        } else {  
            $textBoxes[$i].Text = ""  
            $copyButtons[$i].Enabled = $false  
        }  
        $i++  
    }  
  
    if ($textBoxes[0].Text -ne "") {  
        $findGuidButton.Text = "GUID FOUND"  
        $findGuidButton.Enabled = $false  
    } else {  
        $findGuidButton.Text = "FIND GUID"  
        $findGuidButton.Enabled = $true  
    }  
  
    $progressBar.Value = 100  
    $progressBar.Visible = $false  
})  

$launch_progressBar.Value = 80
$loadingLabel.Text = "Loading interface..."

# Ajouter un gestionnaire pour le changement de texte dans textBoxPath  
$textBoxPath.Add_TextChanged({  
    if (-not $script:fromBrowseButton) {  
        foreach ($textBox in $textBoxes) {  
            $textBox.Text = ""  
        }  
        foreach ($copyButton in $copyButtons) {  
            $copyButton.Enabled = $false  
        }  
        $findGuidButton.Text = "FIND GUID"  
    }  
    $script:fromBrowseButton = $false  
  
    if ($textBoxPath.Text.Trim() -eq "") {  
        $findGuidButton.Enabled = $false  
        $findGuidButton.Text = "FIND GUID"  
    } else {  
        $findGuidButton.Enabled = $true  
    }  
})  
  
if ($textBoxPath.Text.Trim() -eq "") {  
    $findGuidButton.Enabled = $false  
} else {  
    $findGuidButton.Enabled = $true  
}  

$form.Add_Load({
    # This event fires when the form is about to be shown
    $launch_progressBar.Value = 90
    $loadingLabel.Text = "Finalizing..."
    $form.MaximumSize = New-Object System.Drawing.Size([System.Int32]::MaxValue, 375)
    $form.MinimumSize = New-Object System.Drawing.Size(300, 375)

})

$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen  
$form.Add_Resize({ Update-ButtonPositions })  
  
$form.Add_Shown({ 
    $tempPath = "$env:LOCALAPPDATA\Temp\pwsh_loaded.txt" 
    New-Item -Path $tempPath -ItemType File -Force 
    $launch_progressBar.Value = 100
    $loadingLabel.Text = "Complete"
    $loadingForm.Close()
    $form.Activate()
})


[System.Windows.Forms.Application]::Run($form)
