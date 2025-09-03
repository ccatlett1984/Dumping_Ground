# UVtools Script
#Testing for powershell v7, it's required
# Check for required PowerShell version (7+)
if (!($PSVersionTable.PSVersion.Major -ge 7)) {
  try {
    
    # Install PowerShell 7
    if(!(Test-Path "$env:SystemDrive\Program Files\PowerShell\7")) {
      Write-Output 'Installing PowerShell version 7...'
      Invoke-Expression "& { $(Invoke-RestMethod https://aka.ms/install-powershell.ps1) } -UseMSI -Quiet"
    }

    # Refresh PATH
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    
    # Restart script in PowerShell 7
    pwsh "`"$PSCommandPath`"" (($MyInvocation.Line -split '\.ps1[\s\''\"]\s*', 2)[-1]).Split(' ')

  } catch {
    Write-Output 'PowerShell 7 was not installed. Update PowerShell and try again.'
    throw $Error
  } finally { Exit }
}
$inputFile = $null
$slicerFile = $null
#Find the User Documents folder, for default folder when running Get-Filename
$folderpath = "$($env:OneDriveConsumer)\documents"
IF($null -eq (get-childitem -Path $folderpath))
    {
    $folderpath = "$($env:UserProfile)\documents"
    }
Function Get-FolderPath
    {
    param (
    [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName,position=0)]
    [string]$folderpath,
    [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName,position=1)]
    [string]$title
    )
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
    $OpenFolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $OpenFolderDialog.UseDescriptionForTitle = $true
    $OpenFolderDialog.Description = $title
    $OpenFolderDialog.initialDirectory = $folderpath
    $outer = New-Object System.Windows.Forms.Form
    $outer.StartPosition = [Windows.Forms.FormStartPosition] "Manual"
    $outer.Location = New-Object System.Drawing.Point -100, -100
    $outer.Size = New-Object System.Drawing.Size 10, 10
    $outer.add_Shown( { 
    $outer.Activate();
    $outer.DialogResult = $OpenFolderDialog.ShowDialog($outer);
    $outer.Close();
    } )
    $outer.ShowDialog()
    return $OpenFolderDialog.SelectedPath
    } 

Function Get-FileName
    {
    param (
    [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName,position=0)]
    [string]$folderpath,
    [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName,position=1)]
    [string]$title
    )
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = $title
    $OpenFileDialog.initialDirectory = $folderpath
    $outer = New-Object System.Windows.Forms.Form
    $outer.StartPosition = [Windows.Forms.FormStartPosition] "Manual"
    $outer.Location = New-Object System.Drawing.Point -100, -100
    $outer.Size = New-Object System.Drawing.Size 10, 10
    $outer.add_Shown( { 
    $outer.Activate();
    $outer.DialogResult = $OpenFileDialog.ShowDialog($outer);
    $outer.Close();
    } )
    $outer.ShowDialog()
    $OpenFileDialog.FileName
    } 
#end function Get-FileName
#end function Get-FileName

Write-Output "Loading UVTools.Core.dll"
    try
        {
        Add-Type -Path "C:\Program Files\UVtools\UVtools.Core.dll"
        }
    catch
        {
        try
            {
            Add-Type -Path "C:\Program Files (x86)\UVtools\UVtools.Core.dll"
            }
        catch
            {
            Write-Error "Unable to find $coreDll"
            $folderpath = 'C:\Program Files\'
            Add-Type -Path ((Get-Filename $folderpath "Please Select the UVTools.Core.dll file on your system.")[2])
            return
            }
        }  
        

# Progress variable, not really used here but require with some methods
$progress = New-Object UVtools.Core.Operations.OperationProgress

##############
# Dont touch #
##############
# Input file and validation
$files_to_edit = Get-ChildItem -Path ((Get-FolderPath $folderpath "Select the Folder containing sliced files to Edit")[2]) -File 
$dest_folder = ((Get-FolderPath $folderpath "Select the Folder to save Modified Files")[2])
######################################
# All user inputs should be put here #
######################################
# Input iterations number
[decimal]$base_exposure = 0;
while ($base_exposure -le 0) { # Keep asking for a number if the input is invalid
    [decimal]$base_exposure = Read-Host "Enter Burn-in / Base Exposure Value (In Seconds)"
}

[decimal]$normal_exposure = 0;
while ($normal_exposure -le 0) { # Keep asking for a number if the input is invalid
    [decimal]$normal_exposure = Read-Host "Enter Normal Exposure Value (In Seconds)"
}


##############
# Dont touch #
##############
# Decode the file

###################################################
# All operations over the file should be put here #
###################################################
# Morph bottom erode
Write-Output "Updating files with Burn-in Exposure: $("{0:n2}" -f $base_exposure) and Normal Exposure: $("{0:n2}" -f $normal_exposure)"
$count = $files_to_edit.Count
While($count -gt 0)
    {
   
        foreach($file in $files_to_edit)
            {
                try 
                {
                    Write-Output "Decoding file: $($file.Name) Remaining: $($count), please wait..."
                    $slicerFile = [UVtools.Core.FileFormats.FileFormat]::FindByExtensionOrFilePath($File.FullName, $true)
                    $slicerFile.Decode($file.FullName, $progress);
                    $slicerFile.BottomExposureTime = $base_exposure
                    $slicerFile.ExposureTime = $normal_exposure
                    # Save file with _modified name appended
                    $fileExt = [System.IO.Path]::GetExtension($File)
                    $fileOutput = "$($dest_folder)\$($file.BaseName)_$("{0:n2}" -f $normal_exposure)${fileExt}"
                    Write-Output "Saving as $($file.BaseName)_$("{0:n2}" -f $normal_exposure)${fileExt}, please wait..."
                    $slicerFile.SaveAs("$fileOutput", $progress)
                    Write-Output "$fileOutput"
                    Write-Output "Finished!"
                }
                catch
                    {
                    # Catch errors
                    Write-Error "An error occurred:"
                    Write-Error $_.ScriptStackTrace
                    Write-Error $_.Exception.Message
                    Write-Error $_.Exception.ItemName
                    }
            $count--
            }
}