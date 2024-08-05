@echo off & chcp 65001>nul
setlocal enabledelayedexpansion
Title MSI_Properties_Viewer.bat
cd /d "%~dp0"
set "script_path=%~dpnx0"
set "temp_path=%localappdata%\Temp\MSI_Properties_Viewer.bat"
set "spin_counter=0"


for %%I in (A B C D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
    set "script_path=!script_path:%%I=%%I!"
)

set "lower_temp_path=%temp_path%"
for %%I in (A B C D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
    set "lower_temp_path=!lower_temp_path:%%I=%%I!"
)

if "!script_path!" NEQ "!lower_temp_path!" (
    copy "%~dpnx0" "%temp_path%"
    start "" cmd.exe /c "%temp_path%"
    exit
)

call :spin
mode con: cols=55 lines=1


set "psFile=MSI_Properties_Viewer.ps1"

del pwsh_loaded.txt >nul 2>&1
del %psFile% >nul 2>&1
call :spin





echo Add-Type -AssemblyName System.Windows.Forms > %psFile%
echo Add-Type -AssemblyName System.Drawing >> %psFile%
echo.>> %psFile% & call :spin
echo function Get-MsiProperty { >> %psFile%
echo     param ( >> %psFile%
echo         [string]$Path, >> %psFile%
echo         [string]$Property >> %psFile%
echo     ) >> %psFile%
echo     if (Test-Path $Path) { >> %psFile%
echo         $progressBar.Value = 10 >> %psFile%
echo         $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer >> %psFile%
echo         $progressBar.Value = 20 >> %psFile%
echo         $MSIDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($Path, 0)) >> %psFile%
echo         $progressBar.Value = 30 >> %psFile%
echo         $Query = "SELECT Value FROM Property WHERE Property = '$($Property)'" >> %psFile%
echo         $progressBar.Value = 40 >> %psFile%
echo         $View = $MSIDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $MSIDatabase, ($Query)) >> %psFile%
echo         $progressBar.Value = 50 >> %psFile%
echo         $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null) >> %psFile%
echo         $progressBar.Value = 60 >> %psFile%
echo         $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null) >> %psFile%
echo         $progressBar.Value = 70 >> %psFile%
echo         if ($Record) { >> %psFile%
echo             $Value = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1) >> %psFile%
echo             $progressBar.Value = 80 >> %psFile%
echo             $MSIDatabase.GetType().InvokeMember("Commit", "InvokeMethod", $null, $MSIDatabase, $null) >> %psFile%
echo             $progressBar.Value = 90 >> %psFile%
echo             $View.GetType().InvokeMember("Close", "InvokeMethod", $null, $View, $null) >> %psFile%
echo             $MSIDatabase = $null >> %psFile%
echo             $View = $null >> %psFile%
echo             $progressBar.Value = 100 >> %psFile%
echo.>> %psFile% & call :spin
echo             return $Value >> %psFile%
echo         } else { >> %psFile%
echo             return $null >> %psFile%
echo         } >> %psFile%
echo     } else { >> %psFile%
echo         return $null >> %psFile%
echo     } >> %psFile%
echo } >> %psFile%
echo.>> %psFile% & call :spin
echo $form = New-Object System.Windows.Forms.Form >> %psFile%
echo $form.Text = "MSI Properties extractor" >> %psFile%
echo $form.Size = New-Object System.Drawing.Size(600, 375) >> %psFile%
echo.>> %psFile% & call :spin
echo $labels = @() >> %psFile%
echo $textBoxes = @() >> %psFile%
echo $copyButtons = @() >> %psFile%
echo.>> %psFile% & call :spin
echo $properties = [ordered]@{ >> %psFile%
echo     "GUID" = "ProductCode" >> %psFile%
echo     "Product Name" = "ProductName" >> %psFile%
echo     "Product Version" = "ProductVersion" >> %psFile%
echo     "Manufacturer" = "Manufacturer" >> %psFile%
echo     "Upgrade Code" = "UpgradeCode" >> %psFile%
echo } >> %psFile%
echo.>> %psFile% & call :spin
echo # Création des labels, des textboxes et des boutons de copie pour chaque propriété >> %psFile%
echo $yPos = 35 >> %psFile%
echo foreach ($key in $properties.Keys) { >> %psFile%
echo     $label = New-Object System.Windows.Forms.Label >> %psFile%
echo     $label.Text = "$key :" >> %psFile%
echo     $label.AutoSize = $true >> %psFile%
echo     $label.Location = New-Object System.Drawing.Point(10, $yPos) >> %psFile%
echo     $form.Controls.Add($label) >> %psFile%
echo     $labels += $label >> %psFile%
echo.>> %psFile% & call :spin
echo     $textBox = New-Object System.Windows.Forms.TextBox >> %psFile%
echo     $textBox.Size = New-Object System.Drawing.Size(50, 25) >> %psFile%
echo     $textBox.Location = New-Object System.Drawing.Point(100, $yPos) >> %psFile%
echo     $textBox.ReadOnly = $true >> %psFile%
echo     $form.Controls.Add($textBox) >> %psFile%
echo     $textBoxes += $textBox >> %psFile%
echo.>> %psFile% & call :spin
echo     $copyButton = New-Object System.Windows.Forms.Button >> %psFile%
echo     $copyButton.Text = "COPY" >> %psFile%
echo     $copyButton.Size = New-Object System.Drawing.Size(60, 25) >> %psFile%
echo     $copyButton.Enabled = $false >> %psFile%
echo     $form.Controls.Add($copyButton) >> %psFile%
echo     $copyButtons += $copyButton >> %psFile%
echo.>> %psFile% & call :spin
echo     # Ajouter un gestionnaire d'événements pour le bouton de copie >> %psFile%
echo     $copyButton.Add_Click({ >> %psFile%
echo         $buttonIndex = $copyButtons.IndexOf($this) >> %psFile%
echo         [System.Windows.Forms.Clipboard]::SetText($textBoxes[$buttonIndex].Text) >> %psFile%
echo     }) >> %psFile%
echo.>> %psFile% & call :spin
echo     $yPos += 30 >> %psFile%
echo } >> %psFile%
echo.>> %psFile% & call :spin
echo $separatorLine = New-Object System.Windows.Forms.Panel >> %psFile%
echo $separatorLine.Height = 1 >> %psFile%
echo $separatorLine.Width = [int]$form.ClientSize.Width >> %psFile%
echo $separatorLine.BorderStyle = [System.Windows.Forms.BorderStyle]::None >> %psFile%
echo $separatorLine.BackColor = [System.Drawing.Color]::Gray  # Définit la couleur de fond en noir >> %psFile%
echo $separatorLine.AutoSize = $true >> %psFile%
echo $separatorLine.Location = New-Object System.Drawing.Point(0, ($yPos+30)) >> %psFile%
echo $form.Controls.Add($separatorLine) >> %psFile%
echo.>> %psFile% & call :spin
echo # Ajout du label pour MSI Path >> %psFile%
echo $labelMsiPath = New-Object System.Windows.Forms.Label >> %psFile%
echo $labelMsiPath.Text = "MSI Path:" >> %psFile%
echo $labelMsiPath.AutoSize = $true >> %psFile%
echo $labelMsiPath.Location = New-Object System.Drawing.Point(10, $yPos) >> %psFile%
echo $form.Controls.Add($labelMsiPath) >> %psFile%
echo.>> %psFile% & call :spin
echo $yPos += 25  # Augmenter la position Y pour le TextBox du chemin >> %psFile%
echo.>> %psFile% & call :spin
echo $textBoxPath = New-Object System.Windows.Forms.TextBox >> %psFile%
echo $textBoxPath.Size = New-Object System.Drawing.Size(260, 25) >> %psFile%
echo #$textBoxPath.Location = New-Object System.Drawing.Point(10, $yPos) >> %psFile%
echo $form.Controls.Add($textBoxPath) >> %psFile%
echo.>> %psFile% & call :spin
echo $progressBar = New-Object System.Windows.Forms.ProgressBar >> %psFile%
echo $progressBar.Size = New-Object System.Drawing.Size(260, 23) >> %psFile%
echo $progressBar.Style = "Continuous" >> %psFile%
echo $progressBar.Visible = $false >> %psFile%
echo $form.Controls.Add($progressBar) >> %psFile%
echo.>> %psFile% & call :spin
echo $findGuidButton = New-Object System.Windows.Forms.Button >> %psFile%
echo $findGuidButton.Text = "FIND GUID" >> %psFile%
echo $findGuidButton.Size = New-Object System.Drawing.Size(85, 25) >> %psFile%
echo $form.Controls.Add($findGuidButton) >> %psFile%
echo.>> %psFile% & call :spin
echo $browseButton = New-Object System.Windows.Forms.Button >> %psFile%
echo $browseButton.Text = "BROWSE" >> %psFile%
echo $browseButton.Size = New-Object System.Drawing.Size(85, 25) >> %psFile%
echo $form.Controls.Add($browseButton) >> %psFile%
echo.>> %psFile% & call :spin
echo # Variable de contrôle pour différencier l'entrée manuelle et automatique >> %psFile%
echo $script:fromBrowseButton = $false >> %psFile%
echo.>> %psFile% & call :spin
echo # Fonction pour ajuster la position des boutons >> %psFile%
echo function Adjust-ButtonPositions { >> %psFile%
echo     $formWidth = [int]$form.ClientSize.Width >> %psFile%
echo     $formHeight = [int]$form.ClientSize.Height >> %psFile%
echo.>> %psFile% & call :spin
echo     # Ajuster la largeur des TextBox >> %psFile%
echo     $textBoxWidth = $formWidth - 177  # 70 pixels de marge gauche + 40 pixels de marge droite >> %psFile%
echo     foreach ($textBox in $textBoxes) { >> %psFile%
echo         $textBox.Width = $textBoxWidth >> %psFile%
echo     } >> %psFile%
echo     $separatorLine.Width = [int]$form.ClientSize.Width >> %psFile%
echo.>> %psFile% & call :spin
echo     $buttonSpacing = [int](($formWidth - 2 * $findGuidButton.Width) / 3) >> %psFile%
echo     $browseButtonX = $buttonSpacing >> %psFile%
echo     $findGuidButtonX = 2 * $buttonSpacing + $browseButton.Width >> %psFile%
echo     $buttonY = $formHeight - 40 >> %psFile%
echo.>> %psFile% & call :spin
echo     $browseButton.Location = New-Object System.Drawing.Point($browseButtonX, $buttonY) >> %psFile%
echo     $findGuidButton.Location = New-Object System.Drawing.Point($findGuidButtonX, $buttonY) >> %psFile%
echo.>> %psFile% & call :spin
echo     $textBoxPath.Location = New-Object System.Drawing.Point(10, ($buttonY - 30)) >> %psFile%
echo     $textBoxPath.Width = $formWidth - 20 >> %psFile%
echo     $progressBar.Width = $formWidth - 77 >> %psFile%
echo.>> %psFile% & call :spin
echo     # Ajuster la position des boutons de copie >> %psFile%
echo     for ($i = 0; $i -lt $copyButtons.Count; $i++) { >> %psFile%
echo         $copyButtonX = $formWidth - 71 >> %psFile%
echo         $copyButtonY = $textBoxes[$i].Location.Y - 4 >> %psFile%
echo         $copyButtons[$i].Location = New-Object System.Drawing.Point($copyButtonX, $copyButtonY) >> %psFile%
echo     } >> %psFile%
echo.>> %psFile% & call :spin
echo     # Ajuster la position du label MSI Path >> %psFile%
echo     $labelMsiPathY = ($textBoxPath.Location.Y - 25) >> %psFile%
echo     $labelMsiPath.Location = New-Object System.Drawing.Point(10, $labelMsiPathY) >> %psFile%
echo     $progressBar.Location = New-Object System.Drawing.Point(67, ($labelMsiPathY - 3)) >> %psFile%
echo } >> %psFile%
echo.>> %psFile% & call :spin
echo # Appeler la fonction initialement >> %psFile%
echo Adjust-ButtonPositions >> %psFile%
echo.>> %psFile% & call :spin
echo # Ajouter un gestionnaire pour l'événement Resize du formulaire >> %psFile%
echo $form.Add_Resize({ Adjust-ButtonPositions }) >> %psFile%
echo.>> %psFile% & call :spin
echo # Ajouter un gestionnaire pour le bouton Parcourir >> %psFile%
echo $browseButton.Add_Click({ >> %psFile%
echo.>> %psFile% & call :spin
echo     $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog >> %psFile%
echo     $openFileDialog.Filter = "MSI files (*.msi)|*.msi|All files (*.*)|*.*" >> %psFile%
echo     $openFileDialog.Title = "Select a MSI file" >> %psFile%
echo.>> %psFile% & call :spin
echo     if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { >> %psFile%
echo         $selectedFile = $openFileDialog.FileName >> %psFile%
echo         if ($selectedFile -match "\.msi$") { >> %psFile%
echo             if ($selectedFile -ne $textBoxPath.Text) { >> %psFile%
echo                 foreach ($textBox in $textBoxes) { >> %psFile%
echo                     $textBox.Text = "" >> %psFile%
echo                 } >> %psFile%
echo                 foreach ($copyButton in $copyButtons) { >> %psFile%
echo                     $copyButton.Enabled = $false >> %psFile%
echo                 } >> %psFile%
echo                 $progressBar.Visible = $true >> %psFile%
echo                 $progressBar.Value = 5 >> %psFile%
echo                 $job = Start-Job -ScriptBlock { >> %psFile%
echo                     param($path) >> %psFile%
echo                     Start-Sleep -Milliseconds 100 >> %psFile%
echo                     return $path >> %psFile%
echo                 } -ArgumentList $selectedFile >> %psFile%
echo.>> %psFile% & call :spin
echo                 while ($job.State -eq 'Running') { >> %psFile%
echo                     if ($progressBar.Value -lt 90) { >> %psFile%
echo                         $progressBar.Value += 3 >> %psFile%
echo                     } >> %psFile%
echo                     Start-Sleep -Milliseconds 50 >> %psFile%
echo                 } >> %psFile%
echo.>> %psFile% & call :spin
echo                 $result = Receive-Job -Job $job >> %psFile%
echo                 Remove-Job -Job $job >> %psFile%
echo.>> %psFile% & call :spin
echo                 $script:fromBrowseButton = $true >> %psFile%
echo                 $textBoxPath.Text = $result >> %psFile%
echo                 $progressBar.Value = 100 >> %psFile%
echo                 $progressBar.Visible = $false >> %psFile%
echo.>> %psFile% & call :spin
echo                 $findGuidButton.PerformClick() >> %psFile%
echo             } else { >> %psFile%
echo                 [System.Windows.Forms.MessageBox]::Show("Preaching to the choir", "Same File", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) >> %psFile%
echo             } >> %psFile%
echo         } >> %psFile%
echo     } >> %psFile%
echo }) >> %psFile%
echo.>> %psFile% & call :spin
echo # Ajouter un gestionnaire pour le bouton Find GUID >> %psFile%
echo $findGuidButton.Add_Click({ >> %psFile%
echo     $progressBar.Visible = $true >> %psFile%
echo     $progressBar.Value = 0 >> %psFile%
echo     $msiPath = $textBoxPath.Text >> %psFile%
echo     $msiPath = $textBoxPath.Text.Trim('^"') >> %psFile%
echo.>> %psFile% & call :spin
echo     $i = 0 >> %psFile%
echo     foreach ($key in $properties.Keys) { >> %psFile%
echo         $property = $properties[$key] >> %psFile%
echo         $value = Get-MsiProperty -Path $msiPath -Property $property >> %psFile%
echo         if ($value) { >> %psFile%
echo             $textBoxes[$i].Text = $value >> %psFile%
echo             $textBoxes[$i].Text = $textBoxes[$i].Text.Trim() >> %psFile%
echo             $copyButtons[$i].Enabled = $true >> %psFile%
echo         } else { >> %psFile%
echo             $textBoxes[$i].Text = "" >> %psFile%
echo             $copyButtons[$i].Enabled = $false >> %psFile%
echo         } >> %psFile%
echo         $i++ >> %psFile%
echo     } >> %psFile%
echo.>> %psFile% & call :spin
echo     if ($textBoxes[0].Text -ne "") { >> %psFile%
echo         $findGuidButton.Text = "GUID FOUND" >> %psFile%
echo         $findGuidButton.Enabled = $false >> %psFile%
echo     } else { >> %psFile%
echo         $findGuidButton.Text = "FIND GUID" >> %psFile%
echo         $findGuidButton.Enabled = $true >> %psFile%
echo     } >> %psFile%
echo.>> %psFile% & call :spin
echo     $progressBar.Value = 100 >> %psFile%
echo     $progressBar.Visible = $false >> %psFile%
echo }) >> %psFile%
echo.>> %psFile% & call :spin
echo # Ajouter un gestionnaire pour le changement de texte dans textBoxPath >> %psFile%
echo $textBoxPath.Add_TextChanged({ >> %psFile%
echo     if (-not $script:fromBrowseButton) { >> %psFile%
echo         foreach ($textBox in $textBoxes) { >> %psFile%
echo             $textBox.Text = "" >> %psFile%
echo         } >> %psFile%
echo         foreach ($copyButton in $copyButtons) { >> %psFile%
echo             $copyButton.Enabled = $false >> %psFile%
echo         } >> %psFile%
echo         $findGuidButton.Text = "FIND GUID" >> %psFile%
echo     } >> %psFile%
echo     $script:fromBrowseButton = $false >> %psFile%
echo.>> %psFile% & call :spin
echo     # Activer ou désactiver le bouton FIND GUID selon si textBoxPath est vide ou non >> %psFile%
echo     if ($textBoxPath.Text.Trim() -eq "") { >> %psFile%
echo         $findGuidButton.Enabled = $false >> %psFile%
echo         $findGuidButton.Text = "FIND GUID" >> %psFile%
echo     } else { >> %psFile%
echo         $findGuidButton.Enabled = $true >> %psFile%
echo     } >> %psFile%
echo }) >> %psFile%
echo.>> %psFile% & call :spin
echo # Initialiser l'état du bouton FIND GUID selon la valeur initiale de textBoxPath >> %psFile%
echo if ($textBoxPath.Text.Trim() -eq "") { >> %psFile%
echo     $findGuidButton.Enabled = $false >> %psFile%
echo } else { >> %psFile%
echo     $findGuidButton.Enabled = $true >> %psFile%
echo } >> %psFile%
echo.>> %psFile% & call :spin
echo $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen >> %psFile%
echo $form.Add_Load({ $form.MaximumSize = New-Object System.Drawing.Size([System.Int32]::MaxValue, 375); $form.MinimumSize = New-Object System.Drawing.Size(300, 375) }) >> %psFile%
echo $form.Add_Resize({ Adjust-ButtonPositions }) >> %psFile%
echo $tempPath = "$env:LOCALAPPDATA\Temp\pwsh_loaded.txt" >> %psFile%
echo New-Item -Path $tempPath -ItemType File -Force >> %psFile%
echo $form.ShowDialog() >> %psFile%
echo.>> %psFile% & call :spin
echo Start-Sleep -Seconds 1 >> %psFile%
echo Remove-Item -Path $MyInvocation.MyCommand.Definition >> %psFile%
echo.>> %psFile% & call :spin





if exist %psFile% (
    echo PS1 file created with success.
    start "" powershell -WindowStyle Hidden -ExecutionPolicy Bypass -File %psFile%
) else (
    echo Error while creating PS1 file. & pause & exit
)



if not exist pwsh_loaded.txt (
    call :spin
    timeout /t 1 >nul
)


::if not DEFINED IS_MINIMIZED set IS_MINIMIZED=1 && start "" /min "%~dpnx0" %* && exit


del pwsh_loaded.txt >nul 2>&1
START /MIN CMD.EXE /D /C "timeout 1 && del %temp_path%"
exit




:SPIN
set /a spin_counter+=1
if %spin_counter% gtr 13 set "spin_counter=1"
set "LINE=< LOADING "
for /L %%C in (1,1,%spin_counter%) do set "LINE=!LINE!^>"
title !LINE!
goto :EOF