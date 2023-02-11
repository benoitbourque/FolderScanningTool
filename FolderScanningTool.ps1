<#
.SYNOPSIS
    Folder Scanning Tool
.DESCRIPTION
    Scan recursif du contenu du dossier sélectionné.
    Output le resulat du scan dans un GridView en fonction des threshold configurés.
    Actions possibles avec la sélection:
        - Affichage de la liste sélectionnés
        - Modification de la sélection.
        - Suppression des items sélectionnés.
        - Exportation de la liste d'items sélectionnés en CSV.
    Il est aussi possible de relancer un scan a partir la console.
    Le threshold pour la taille des dossiers est en GB.
    Le threshold pour la taille des fichiers est en MB.

.NOTES
    Fichier    : FolderScanningTool_V2.ps1
    Author     : Benoit Bourque - benoit.bourque@cpu.ca
    Version    : 1.2
    Updated    : 2023-02-11
#>

#Function Get-BaseFolder
#---------------------------------------------------------------
<# File Browser Dialog
    Let script user select a basefolder.
#>
Function Get-BaseFolder
{
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dialog.Description = 'Select folder to scan'
    $dialog.RootFolder = [System.Environment+specialfolder]::MyComputer 
    $dialog.ShowNewFolderButton = $true
    $dialog.ShowDialog()
    $PathtextBox.Clear()
    $PathtextBox.text = $dialog.SelectedPath
}
 
#Function Add-ToArray
#---------------------------------------------------------------
<# Adds files/folders larger than threshold to array.
    Array is used to generate gridview output
    Params :
        Value      = Path
        Size(MB)   = Folder/file size in MB
        Size(GB)   = Folder/file size in GB
        Type       = File or Folder
        Creation   = Creation time
        LastWrite  = Last write time
        LastAccess = Last Access time
        Reason     = Reason

#>   
function Add-ToArray
    {
    #Params
    param($Value,$ValueSizeinMB,$ValueSizeInGB,$Type,$LastWrite,$LastAccess,$Creation,$Reason)
    #SizeVars
    $ValueSizeInGB = [FLOAT]$ValueSizeInGB
    $ValueSizeinMB = [FLOAT]$ValueSizeinMB
    #PSCustomObject
    $ValueObject = [PSCustomObject]@{
        Path    = [STRING]$Value.FullName
        'Size(MB)'    = $ValueSizeInMB
        'Size(GB)'    = $ValueSizeinGB
        'Type'        = $Type
        'LastAccess'  = $LastAccess
        'LastWrite'   = $LastWrite
        'Creation'    = $Creation
        'Reason'      = $Reason
    }   
    #Add object to array
    $Array.Add($ValueObject) | Out-Null
    
    }

#Function Add-ToArray
#---------------------------------------------------------------
<# Add selection to exclusions list
#>
function Add-Exclusions
    {
    $NewList += @(Get-Content .\exclusions.txt) 
    ForEach($SelectedItem in $SelectedItems)
        {
        $NewList += $SelectedItem.Path
        Write-Host "[INFO] - Adding" $SelectedItem.Path "to exclusions list." -ForegroundColor Yellow                 
        }
    Write-Host "[INFO] - Writing changes to exclusions list file." -ForegroundColor Green
    Set-Content -Path .\exclusions.txt -Value $NewList                  
    pause
    }

#Function Browse-Folder
#---------------------------------------------------------------
<# Browse selected folder in explorer.exe
#>
function Browse-Folder
    {
    $Folders = $SelectedItems | Where-Object -Property Type -EQ Folder
    $menu = @{}
    for ($i=1;$i -le $Folders.path.count; $i++) 
    { Write-Host "$i. $($Folders[$i-1].Path)"
    $menu.Add($i,($Folders[$i-1].Path)) }

    [int]$ans = Read-Host 'Enter selection'
    $Browse = $menu.Item($ans) ; explorer $Browse
    }
#Function Scan-ChildFiles
#---------------------------------------------------------------
<# Scan files in folders larger than threshold(GB):
# If file size is larger than threshold(MB):
#    - Add file to array
#>
Function Scan-ChildFiles
    {
    #Param
    param($Folder)
    #Get files in folder
    $ChildFilesQuery = Get-ChildItem $Folder.Fullname | Select-Object -Property Mode,FullName,Length,LastWriteTime,LastAccessTime,CreationTime | Where-Object -Property Mode -like "*a*"
    
    #Exclusions   
    $ChildFiles = @()
    $ChildFiles = $ChildFilesQuery.Where({$path = $_.fullname; -not $Exclusions.Where({ $path -eq "$_" })})   
    #Scan files
    foreach($ChildFile in $ChildFiles)
        {
        #Console output
        $target = $ChildFile.Fullname
        write-host "[INFO] - Scanning $target..." -ForegroundColor White
        #Get file size
        $FileSize = $(Get-Childitem -Path $target -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).sum
        #SizeVars
        [FLOAT]$FileSizeInGB = "{0:N2}" -f $($FileSize / 1Gb)
        [FLOAT]$FileSizeinMB = "{0:N2}" -f $($FileSize / 1MB)
        #If file is larger than threshold(MB)        
        if ($FileSizeinMB -gt $FilesizeThresholdInMb)
            {
            #Reason
            $Reason = "File is larger than threshold [$FilesizeThresholdInMb MB]."
            #Add to array
            Add-ToArray -Value $ChildFile -ValueSizeinGB $FileSizeInGB -ValueSizeinMB $FileSizeinMB -type "File" -LastAccess $ChildFile.LastAccessTime  -LastWrite $ChildFile.LastWriteTime -Creation $ChildFile.CreationTime -Reason $Reason
            }
        }
    }

#Function Scan-Folders
#---------------------------------------------------------------
<# Scan selected folders and sub-folders.
# If folder size is larger than threshold:
#    - Add folder to array
#    - Scan for files larger than threshold
#    - Files larger than threshold are added to array
#>
Function Scan-Folders
    {
    #Get files located in selected folder's root
    $Root = Get-Item $Global:DefaultPath | Select-Object -Property Mode,FullName,Length,LastWriteTime,LastAccessTime,Creationtime
    #Get sub-folders
    $RootFoldersQuery = Get-ChildItem $Global:DefaultPath  -Directory -Recurse | Select-Object -Property Mode,FullName,Length,LastWriteTime,LastAccessTime,Creationtime | Where-Object -Property Mode -like "d*"
    #Exclusions
    $Exclusions = @(Get-Content .\exclusions.txt | Select-Object -Skip 1)
    ForEach($Exclusion in $Exclusions)
        {
        write-host "[WARNING] - $Exclusion excluded from scan..." -ForegroundColor Yellow
        }    
    #Folder to scan 
    $RootFolders = @()
    $RootFolders = $RootFoldersQuery.Where({$path = $_.fullname; -not $Exclusions.Where({ $path -like "*$_*" })})   
    #Scan files in selected folder's root 
    Scan-ChildFiles -Folder $Root
    #Scan sub-folders
    ForEach($RootFolder in $RootFolders)
        {
        #Console output
        $target = $RootFolder.Fullname    
        write-host "[INFO] - Scanning $target..." -ForegroundColor White
        #Get folder size
        $FolderSize = $(Get-Childitem -Path $target -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).sum
        $filescount = $(Get-Childitem -Path $target | Where-Object -Property Mode -like "-a*").count
        #SizeVars
        [FLOAT]$FolderSizeInGB = "{0:N2}" -f $($FolderSize / 1Gb)
        [FLOAT]$FolderSizeinMB = "{0:N2}" -f $($FolderSize / 1MB)
        #If folder is larger than threshold(GB)
        if ($FolderSizeInGB -gt $FolderSizeThresholdInGB -and $filescount -gt $FilesCountThreshold )
            {
            #Reason
            $Reason = "Folder size is larger than threshold [$FolderSizeThresholdInGB GB]. File count is higher than threshold [$FilesCountThreshold]."
            #Add to array
            Add-ToArray -Value $RootFolder -ValueSizeInGB $FolderSizeInGB -ValueSizeinMB $FolderSizeinMB -type "Folder" -LastWrite $RootFolder.LastWriteTime -LastAccess $rootFolder.LastAccessTime -Creation $RootFolder.CreationTime -Reason $Reason 
            #Scan files in sub-folder
            Scan-ChildFiles -Folder $RootFolder
            }
        elseif ($FolderSizeInGB -gt $FolderSizeThresholdInGB)
            {
            #Reason
            $Reason = "Folder size is larger than threshold [$FolderSizeThresholdInGB GB]."
            #Add to array
            Add-ToArray -Value $RootFolder -ValueSizeInGB $FolderSizeInGB -ValueSizeinMB $FolderSizeinMB -type "Folder" -LastWrite $RootFolder.LastWriteTime -LastAccess $rootFolder.LastAccessTime -Creation $RootFolder.CreationTime -Reason $Reason
            #Scan files in sub-folder
            Scan-ChildFiles -Folder $RootFolder
            }
        elseif ($filescount -gt $FilesCountThreshold)
            {
            #Reason
            $Reason = "File count is higher than threshold [$FilesCountThreshold]."
            #Add to array
            Add-ToArray -Value $RootFolder -ValueSizeInGB $FolderSizeInGB -ValueSizeinMB $FolderSizeinMB -type "Folder" -LastWrite $RootFolder.LastWriteTime -LastAccess $rootFolder.LastAccessTime -Creation $RootFolder.CreationTime -Reason $Reason
            #Scan files in sub-folder
            Scan-ChildFiles -Folder $RootFolder
            }
        } 
    }

#Function Export-To-CSV
#---------------------------------------------------------------
<# Export selected items properties to CSV.
#>
Function Export-To-CSV
    {
    #CSVDATA Array
    $CSVDATA = @()
    #GetDate
    $Date = $(Get-Date -Format "MM-dd-yyyy_HHmmss")
    $CSVFile = "ScanResult_" + $Date + ".csv"
    #Add data to CSVDATA Array
    Foreach($SelectedFile in $SelectedItems)
        {
        #VarsDeclaration
        $Type = $SelectedFile.Type
        $Path = $SelectedFile.Path
        $SizeGB = $SelectedFile | Select-Object -ExpandProperty "Size(GB)"
        $SizeMB = $SelectedFile | Select-Object -ExpandProperty "Size(MB)"
        $Creation = $SelectedFile.Creation
        $LastWrite = $SelectedFile.LastWrite
        $LastAccess = $SelectedFile.LastAccess
        $Reason = $SelectedFile.Reason 

        #PSObject
        $row = New-Object PSObject
        $row | Add-Member -MemberType NoteProperty -Name "Type" -Value $Type
        $row | Add-Member -MemberType NoteProperty -Name "Path" -Value $Path
        $row | Add-Member -MemberType NoteProperty -Name "Size(GB)" -Value $SizeGB
        $row | Add-Member -MemberType NoteProperty -Name "Size(MB)" -Value $SizeMB
        $row | Add-Member -MemberType NoteProperty -Name "Creation" -Value $Creation
        $row | Add-Member -MemberType NoteProperty -Name "Last Write" -Value $LastWrite
        $row | Add-Member -MemberType NoteProperty -Name "Last Access" -Value $LastAccess
        $row | Add-Member -MemberType NoteProperty -Name "Reason" -Value $Reason
        #Add data to CSVDATA Array
        $CSVDATA += $row
        }
    #Export CSVDATA Array to CSV file
    $CSVDATA | Sort-Object -Property "Size(GB)" | Export-Csv $CSVFile -NoTypeInformation
    }

#Function Remove-Selection
#---------------------------------------------------------------
<# Remove selected items. Folders are recursively removed.
#>
function Remove-Selection
    {
    #FILES
    write-host "============ Removing Files ============"
    $FilesToRemove = $SelectedItems | Where-Object -Property Type -EQ File
    ForEach($FileToRemove in $FilesToRemove)
        {                   
        try #Remove files
            {
            Remove-Item $FileToRemove.Path -Force -ErrorAction SilentlyContinue
            }
        Catch #Catch Error
            {
            $Err = $_
            }
        Finally #Check
            {
            if(Test-Path -Path $FileToRemove.Path)
                {
                Write-host "[ERROR] - Could not remove file :" $Err -ForegroundColor Red 
                }
            else
                {
                write-host "[INFO] -" $FileToRemove.Path "removed" -ForegroundColor Green 
                }
            }                                  
        }
    #FOLDERS
    write-host "============ Removing Folders ============"
    $FoldersToRemove = $SelectedItems | Where-Object -Property Type -EQ Folder
    ForEach($FolderToRemove in $FoldersToRemove)
        { 
        try #Remove folders
            {
            Remove-Item $FolderToRemove.Path -Force -Recurse -ErrorAction SilentlyContinue
            }
        Catch #Catch Error
            {
            $Err = $_

            }
        Finally #Check
            {
            if(Test-Path -Path $FolderToRemove.Path)
                {
                Write-host "[ERROR] - Could not remove folder :" $Err -ForegroundColor Red 
                }
            else
                {
                write-host "[INFO] -" $FolderToRemove.Path "removed" -ForegroundColor Green 
                }
            }
        } 
    }

#Function Show-Selection
#---------------------------------------------------------------
<# Display list of selected items in console
#>
function Show-Selection
    {
    write-host "============ Selection ============"
    ForEach($SelectedFile in $SelectedItems)
        {                 
        write-host $SelectedFile.Path
        }
    }

function Check-Selection
    {
    If($SelectedItems)
        {$Global:Array | Sort-Object {$_.Filename }
        $TotalSizeInGB = 0
        $TotalSizeInGB = "{0:N2}" -f $($SelectedItems | Select-Object -ExpandProperty "Size(GB)"  |Measure-Object -sum ).sum
        }
    else
        {
        $SelectedItems = $Global:Array | Sort-Object {$_.Filename }
        $TotalSizeInGB = 0
        $TotalSizeInGB = "{0:N2}" -f $($SelectedItems | Select-Object -ExpandProperty "Size(GB)"  |Measure-Object -sum ).sum
        }
    }

#Function Show-Selection
#---------------------------------------------------------------
<# Open gridview to modify selection
#>
function Modify-Selection
    {
    $SelectedItems = ""
    $SelectedItems = $Global:Array | Sort-Object {$_.Filename }| Out-GridView -Title "Please make a selection" –PassThru
    Check-Selection
    }



#Function Show-Menu
#---------------------------------------------------------------
<# Generate menu
#>
function Show-Menu {
    param (
        [string]$Title = 'Actions'
    )
    Clear-Host
    $ItemsCount = $($SelectedItems.Path).count
    $TotalSizeInGB = 0
    $TotalSizeInGB = "{0:N2}" -f $($SelectedItems | Select-Object -ExpandProperty "Size(GB)"  |Measure-Object -sum ).sum
    Write-Host "================ Summary ================"
    Write-Host "Items Selected :" $ItemsCount
    Write-Host "Total size in GB :" $TotalSizeInGB "GB"
    Write-Host "================ $Title ================"
    
    Write-Host "V: Press 'V' to view selection."
    Write-Host "B: Press 'B' to browse a folder."
    Write-Host "M: Press 'M' to modify selection."
    Write-Host "D: Press 'D' to delete ALL selected items."
    Write-Host "E: Press 'E' to export selection to CSV."
    Write-Host "R: Press 'R' to rescan."
    Write-Host "X: Press 'X' to add selection to exclusions."
    Write-Host "Q: Press 'Q' to quit."
}

#Function Show-ActionMenu
#---------------------------------------------------------------
<# Show actions menu
#>
function Show-ActionMenu
    {
    do
        {
        Show-Menu
        $selection = Read-Host "Please make a selection"
        #Actions switch
        switch ($selection)
            {
            'V' #Display list of selected items in console
                {
                Show-Selection
                pause
                }
            'B' #Browse selected folder
                {
                Browse-Folder
                }
            'M' #Modify selection
                {
                Modify-Selection
                }
            'D' # Remove selected items. Folders are recursively removed.
                {
                Remove-Selection
                pause
                }
            'E' #Export selected items properties to CSV.
                {
                Export-To-CSV
                }
            'R' #Rescan
                {
                Clear-Host
                FolderScanningTool
                }
            'x' #Add to exclude list
                { 
                Add-Exclusions
                }
            }

        }
    #Quit
    until ($selection -eq 'q')
    }


#---------------------------------------------------------------
<# FolderScanningTool
#>
Function FolderScanningTool
    {
    #Array
    [System.Collections.ArrayList]$Global:Array = @()
    #Default path - C:\users\
    $Global:DefaultPath = $(Set-Location $env:USERPROFILE ; Set-Location ..\ ; Get-Location | Select-Object -ExpandProperty Path)

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    #FORM
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Configuration'
    $form.Size = New-Object System.Drawing.Size(350,250)
    $form.StartPosition = 'CenterScreen'
    $form.Topmost = $true

    #OKBUTTON
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(150,175)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    #CANCELBUTTON
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(225,175)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)

    #FOLDERSIZELABEL
    $FolderSizelabel = New-Object System.Windows.Forms.Label
    $FolderSizelabel.Location = New-Object System.Drawing.Point(10,20)
    $FolderSizelabel.Size = New-Object System.Drawing.Size(265,20)
    $FolderSizelabel.Text = 'Minimum folder size in GB (Default = 5):'
    $form.Controls.Add($FolderSizelabel)

    #FOLDERSIZETEXTBOX
    $FolderSizetextBox = New-Object System.Windows.Forms.TextBox
    $FolderSizetextBox.Location = New-Object System.Drawing.Point(280,20)
    $FolderSizetextBox.Size = New-Object System.Drawing.Size(40,20)
    $FolderSizetextBox.Text = 5
    $FolderSizetextBox.TextAlign = "Right"
    $form.Controls.Add($FolderSizetextBox)

    #FILESCOUNTLABEL
    $FilesCountlabel = New-Object System.Windows.Forms.Label
    $FilesCountlabel.Location = New-Object System.Drawing.Point(10,50)
    $FilesCountlabel.Size = New-Object System.Drawing.Size(260,20)
    $FilesCountlabel.Text = 'Files count threshold (Default = 25): '
    $form.Controls.Add($FilesCountlabel)
    
    #FILESCOUNTEXTBOX
    $FilesCounttextBox = New-Object System.Windows.Forms.TextBox
    $FilesCounttextBox.Location = New-Object System.Drawing.Point(280,50)
    $FilesCounttextBox.Size = New-Object System.Drawing.Size(40,20)
    $FilesCounttextBox.Text = 25
    $FilesCounttextBox.TextAlign = "Right"
    $form.Controls.Add($FilesCounttextBox)

    #FILESIZELABEL
    $FileSizelabel = New-Object System.Windows.Forms.Label
    $FileSizelabel.Location = New-Object System.Drawing.Point(10,80)
    $FileSizelabel.Size = New-Object System.Drawing.Size(260,20)
    $FileSizelabel.Text = 'Minimum file size in MB (Default = 150):'
    $form.Controls.Add($FileSizelabel)
    
    #FILESIZETEXTBOX
    $FileSizetextBox = New-Object System.Windows.Forms.TextBox
    $FileSizetextBox.Location = New-Object System.Drawing.Point(280,80)
    $FileSizetextBox.Size = New-Object System.Drawing.Size(40,20)
    $FileSizetextBox.Text = 150
    $FileSizetextBox.TextAlign = "Right"
    $form.Controls.Add($FileSizetextBox)



    #PATHBUTTON
    $PathButton = New-Object System.Windows.Forms.Button
    $PathButton.Location = New-Object System.Drawing.Point(75,175)
    $PathButton.Size = New-Object System.Drawing.Size(75,23)
    $PathButton.Text = 'Path'
    $form.Controls.Add($PathButton)

    #PATHTEXTBOT
    $PathtextBox = New-Object System.Windows.Forms.TextBox
    $PathtextBox.Location = New-Object System.Drawing.Point(10,120)
    $PathtextBox.Size = New-Object System.Drawing.Size(310,20)
    $PathtextBox.text = $Global:DefaultPath
    $PathtextBox.TextAlign = "Right"
    $PathtextBox.Enabled   = $False
    $form.Controls.Add($PathtextBox)

    #PATHBUTTON - ADD CLICK : File Browser Dialog
    $PathButton.add_click({ 
        Get-BaseFolder
        })
    
    #Set to location to Scriptroot
    $ScriptRoot = $(Set-Location $PSScriptRoot ; Get-Location | Select-Object -ExpandProperty Path)
    Set-Location $ScriptRoot

    #SHOWDIALOG
    $result = $form.ShowDialog()

    #IF OK BUTTON
    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
        {
        $FolderSizeThresholdInGB = $FolderSizetextBox.text
        $FilesizeThresholdInMb   = $FileSizetextBox.text
        $FilesCountThreshold = $FilesCounttextBox.Text
        $Global:DefaultPath = $PathtextBox.text
        Scan-Folders
        $SelectedItems = $Global:Array | Sort-Object {$_.Filename }| Out-GridView -Title "Please make a selection" –PassThru
        Check-Selection
        Show-ActionMenu
        }
    
    #IF CANCEL BUTTON
    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
        {
        Exit
        }
    }

#START FOLDER SCANNING TOOL
#---------------------------------------------------------------
FolderScanningTool




