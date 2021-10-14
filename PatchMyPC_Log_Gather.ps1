$log_bundle_path = $($env:temp + "\PMPC_log_Bundle $(get-date -Format "MM_dd_yyyy").zip")

Get-WindowsUpdateLog -Confirm:$false | out-null

$PMPC_logfiles = $($env:WinDir + "\CCM\Logs\CAS*.log"),
$($env:WinDir + "\CCM\Logs\CIAgent.*log"),
$($env:WinDir + "\CCM\Logs\ClientLocation*.log"),
$($env:WinDir + "\CCM\Logs\CMBITSManager*.log"),
$($env:WinDir + "\CCM\Logs\ContentTransferManager*.log"),
$($env:WinDir + "\CCM\Logs\DataTransferService*.log"),
$($env:WinDir + "\CCM\Logs\StateMessage.log"),
$($env:WinDir + "\CCM\Logs\LocationServices.log*.log"),
$($env:WinDir + "\CCM\Logs\UpdatesDeployment*.log"),
$($env:WinDir + "\CCM\Logs\UpdatesHandler*.log"),
$($env:WinDir + "\CCM\Logs\UpdatesStore*.log"),
$($env:WinDir + "\CCM\Logs\PatchMyPC-ScriptRunner.log"),
$($env:WinDir + "\CCM\Logs\CAS*.log"),
$($env:WinDir + "\CCM\Logs\DeltaDownload*.log"),
$($env:WinDir + "\CCM\Logs\DataTransferService*.log"),
$($env:WinDir + "\CCM\Logs\PatchMyPC-ScriptRunner.log (If exist)"),
$($env:WinDir + "\CCM\Logs\ScanAgent*.log"),
$($env:WinDir + "\CCM\Logs\StateMessage.log"),
$($env:WinDir + "\CCM\Logs\UpdatesDeployment*.log"),
$($env:WinDir + "\CCM\Logs\UpdatesHandler*.log"),
$($env:WinDir + "\CCM\Logs\UpdatesStore*.log"),
$($env:WinDir + "\CCM\Logs\WUAHandler*.log"),
$($env:WinDir + "\WindowsUpdate.log"),
$($env:ProgramData + "\PatchMyPC\PatchMyPC-UserNotification.log"),
$($env:WinDir + "\CCM\Logs\AppDiscovery*.log"),
$($env:WinDir + "\CCM\Logs\AppEnforce*.log"),
$($env:WinDir + "\CCM\Logs\AppIntentEval*.log"),
$($env:WinDir + "\CCM\Logs\CAS*.log"),
$($env:WinDir + "\CCM\Logs\CIAgent.*log"),
$($env:WinDir + "\CCM\Logs\DataTransferService*.log"),
$($env:WinDir + "\CCM\Logs\PatchMyPC-ScriptRunner.log"),
$($env:WinDir + "\CCM\Logs\PatchMyPC-SoftwareDetectionScript.log"),
$($env:WinDir + "\CCM\Logs\StateMessage.log"),
$($env:ProgramData + "\PatchMyPC\PatchMyPC-UserNotification.log"),
$($env:ProgramData + "\PatchMyPCIntuneLogs\PatchMyPC-ScriptRunner.log"),
$($env:ProgramData + "\PatchMyPCIntuneLogs\PatchMyPC-SoftwareDetectionScript.log"),
$($env:ProgramData + "\PatchMyPCIntuneLogs\PatchMyPC-SoftwareUpdateDetectionScript.log"),
$($env:ProgramData + "\Microsoft\IntuneManagementExtension\Logs\AgentExecutor.log"),
$($env:ProgramData + "\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log"),
$($env:ProgramData + "\PatchMyPC\PatchMyPC-UserNotification.log")

foreach ($log in $PMPC_logfiles)
    {
        #get files that match wildcards.
        $files = $null
        $files = Get-ChildItem -Path ($log | Split-Path) -Filter "*.$(($log.Split(".",2))[1])" | Where-Object -Property "Name" -Like "*$((($log | split-path -Leaf).Split(".",2))[0])"
        foreach($file in $files)
            {
            #check that the file still exists.
            if(test-path -Path $file)
                {
                    try 
                        {
                        #archive the file
                        $file | compress-archive -DestinationPath $log_bundle_path -update -verbose
                        }
                    catch 
                        {
                        #file was locked, make a copy, then archive that.....
                        copy-item -path $file -destination $($env:temp + "\$($file.name)") -force
                        $file | compress-archive -DestinationPath $log_bundle_path -update
                        remove-item -path $($env:temp + "\$($file.name)") -force
                        }
                }
            else
                {
                #oops, its gone....
                write-warning "File Not Found: $file"
                }
            }       
    }
