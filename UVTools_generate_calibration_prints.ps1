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
$open_count = 3
$inputFile = $null
while (($null -eq $inputFile) -and ($open_count -gt 0))
    { # Keep asking for a file if the input is invalid
    write-output "Please select the Base Slicer File in the dialog box"
    $inputFile = (Get-FileName $folderpath "Please Select the Base Slicer File")[2]
        if($null -eq $inputFile)
            {
            return;
            }
        if((get-item $inputFile).psiscontainer)
            {
            Write-host "Input file must be an valid file, re-enter."
            $inputFile = $null
            }
        else
            {
            $slicerFile = [UVtools.Core.FileFormats.FileFormat]::FindByExtensionOrFilePath($inputFile, $true)
            if(!$slicerFile)
                {
                Write-host "Invalid file format, re-enter."
                $inputFile = $null
                }
            }
        $open_count--
    }

IF($open_count -eq 0){
    exit
}

######################################
# All user inputs should be put here #
######################################
# Input iterations number
[decimal]$max = 0;
while ($max -le 0) { # Keep asking for a number if the input is invalid
    [decimal]$max = Read-Host "Enter Maximum Exposure Value (In Seconds)"
}
IF($max.ToString() -notlike "*.*")
    {

    }

[decimal]$step_value = 0;
while ($step_value -le 0) { # Keep asking for a number if the input is invalid
    [decimal]$step_value = Read-Host "Enter Value for each Exposure Step"
}

[int]$iterations = 0;
while ($iterations -le 0) { # Keep asking for a number if the input is invalid
    [int]$iterations = Read-Host "Number of Exposure Step iterations"
}

##############
# Dont touch #
##############
# Decode the file
Write-Output "Decoding, please wait..."
$slicerFile.Decode($inputFile, $progress);


###################################################
# All operations over the file should be put here #
###################################################
# Morph bottom erode
Write-Output "Creating $iterations files with Max Exposure: $("{0:n2}" -f $max) and a Step Exposure of $step_value"
$count = $iterations
While($count -gt 0)
    {
    try 
        {   
        $exposure = $max
        $slicerFile.ExposureTime = $exposure
       
        # Save file with _modified name appended
        $filePath = [System.IO.Path]::GetDirectoryName($inputFile)
        $fileExt = [System.IO.Path]::GetExtension($inputFile)
        $fileNoExt = [System.IO.Path]::GetFileNameWithoutExtension($inputFile)
        $fileOutput = "${filePath}\${fileNoExt}_$("{0:n2}" -f $exposure)${fileExt}"
        Write-Output "Saving as ${fileNoExt}_$("{0:n2}" -f $exposure)${fileExt}, please wait..."
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
$max = $max - $step_value
$count--
}