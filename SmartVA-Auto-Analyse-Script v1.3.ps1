#Load Parameters from configuration file
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
Get-Content "config.txt" | foreach-object -begin {$config=@{}} -process { $k = [regex]::split($_,'='); if(($k[0].CompareTo("") -ne 0) -and ($k[0].StartsWith("[") -ne $True)) { $config.Add($k[0], $k[1]) } }

#Variables
$counter = 1
$errors = $null
$errorcount = 0
$allresults= $null
$output = $null
$resultcount = 0
$archiveResultErrorCount = 0
$archiveErrorMsg = $null
$object = $null
$logfiledir = $config.ProcessDir + "\Logs"
$logfile = $logfiledir + "\log.txt"
$collectDir = $config.ProcessDir + "\CollectFiles"
$collectFormsDir = $collectDir + "\" + $config.FormsDir
$collectInstancesDir = $collectDir + "\" + $config.InstancesDir
$collectArchiveDir = $config.ArchiveDir + "\Archive"
$processedArchiveFormsDir = $collectArchiveDir + "\Processed\forms"
$processedArchiveInstancesDir = $collectArchiveDir + "\Processed\instances"
$errorsArchiveFormsDir = $collectArchiveDir + "\Errors\forms"
$errorsArchiveInstancesDir = $collectArchiveDir + "\Errors\instances"
$resultsfile = $collectArchiveDir + "\results.csv"
$mccod = $collectArchiveDir + "\mccod.csv"
$gridfile = $logfiledir + "\gridview.csv"
$odkbcdir = $config.ProcessDir + "\ODK Briefcase Storage"
$odkdir = $config.ODKDir
$briefcaseFormID = $config.FormID
$briefcaseStorageDir = $config.ProcessDir
$briefcaseExportDir = $odkbcdir + "\Export"
$briefcaseConvertedFile = "TariffReadyFile.csv"
$tariffOutput = $config.ProcessDir + "\TariffFiles"
$tariffCountry = $config.country
$tariffHIV = $config.hiv
$tariffMalaria = $config.malaria
$tariffHCE = $config.hce
$tariffFreetext = $config.freetext
$tariffFigures = $config.figures
$tariffInput = $briefcaseExportDir + "\" + $briefcaseConvertedFile
$tariffError = $null
$PhdUCod = $null

#Check for Processing Directory and if it does not exist, create it and associated sub-directories
if(!(Test-Path $config.ProcessDir)){
    New-Item $config.ProcessDir -ItemType directory
}
if(!(Test-Path $logfiledir)){
    New-Item $logfiledir -ItemType directory
}
if(!(Test-Path $collectDir)){
    New-Item $collectDir -ItemType directory
}
if(!(Test-Path $processedArchiveFormsDir)){
    New-Item $processedArchiveFormsDir -ItemType directory
}
if(!(Test-Path $processedArchiveInstancesDir)){
    New-Item $processedArchiveInstancesDir -ItemType directory
}
if(!(Test-Path $resultsfile)){
    #create results file and add headers
    "VA_ID," + "Age," + "Sex," + "CoD1," + "Lh1," + "CoD2," + "Lh2," + "CoD3," + "Lh3," + "Date_Entered" | Add-Content $resultsfile
}
if(!(Test-Path $mccod)){
    #create results file and add headers
    "VA_ID," + "CoD1," + "mccod" | Add-Content $mccod
}

#Remove Temp File
if(Test-Path $gridfile){
   Remove-Item $gridfile
}

#Add Date/Time stamp of current run to log file and create datetime2 variable, formatted for use with Archive Directory Creation
"`r" | Add-Content $logfile
$datetime = Get-Date
$datetime | Add-Content $logfile
$datetime2 = "{0:yyyyMMdd_hh_mm_ss}" -f (get-date)


#If Android 2.0 with USB Mass Storage capability - set AndroidOld setting in Config.txt to 1

if($config.AndroidOld -eq 1){
    
    #Get Drive Letters of all USB attached devices (used for older Android 2.0 devices with USB Mass Storage capability)
    $diskdrive = gwmi win32_diskdrive | ?{$_.interfacetype -eq "USB"}
    $letters = $diskdrive | %{gwmi -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID=`"$($_.DeviceID.replace('\','\\'))`"} WHERE AssocClass = Win32_DiskDriveToDiskPartition"} |  %{gwmi -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID=`"$($_.DeviceID)`"} WHERE AssocClass = Win32_LogicalDiskToPartition"} | %{$_.deviceid}


    #Check if USB attached (used for older Android 2.0 devices with USB Mass Storage capability)
    if ([string]::IsNullOrEmpty($letters)){
        (new-object -ComObject wscript.shell).Popup("No USB Drive attached`n`nCheck that you have a USB Drive connected.  If using a tablet with Android 4.x or above installed, check that the setting 'AndroidOld' is set to '0' in the config.txt file",0,"Error!")
        "No USB Drive attached" | Add-Content $logfile
        exit 
    }else{

        #create array of drive letters found
        $usbattached = gwmi win32_volume | ? {$letters -contains ($_.name -replace "\\")} | select name

        #check for odk folder on USB device
        ForEach ($device in $usbattached){
            $letter = $($device.name)
            $odkpath = $letter + $config.ODKDir
            if(Test-Path $odkpath){
                $pathfound = "1"
                break
            }
        }

        if($pathfound -eq 1){
        
            "ODK Folder found at $odkpath`r" | Add-Content $logfile
            
            #Lookup folder/directory names for all completed forms on the Tablet
            $odkdirs = Get-ChildItem -Path $odkpath -Recurse -include forms,instances | Where-Object { $_.PSIsContainer }
            
            #Check if the Tablet form files already exist in Processing\CollectFiles Directory, delete previous files if found, and if not found, copy across for further processing
             ForEach ($dir in $odkdirs){
                $dirname = $collectDir + "\" + $dir.Name
                if((Test-Path $dirname)){
                    Remove-Item $dirname -Force -Recurse
                    Copy-Item -Path $dir.Fullname -Destination $collectDir -Recurse -Container
                    "Directory $dirName copied from device`r" | Add-Content $logfile
                }else{
                    Copy-Item -Path $dir.Fullname -Destination $collectDir -Recurse -Container
                    "Directory $dirName copied from device`r" | Add-Content $logfile
                }
            }
        }else{
            "ODK folder not found on attached USB Device.  Check config.txt file has correct location of ODK folder set and ODK Collect is installed on the device" | Add-Content $logfile
            (new-object -ComObject wscript.shell).Popup("ODK folder not found on attached USB Device. Check config.txt file has correct location of ODK folder set and ODK Collect is installed on the device",0,"Error!")
            exit
        }
    }
}else{

    #Check if Android Device is connected and in device mode ready for copy.  If not, throw error, if ready, copy files from ODK location to local PC
    if((adb ($config.DeviceType) get-state) -eq "device"){
        "Android Device found....`r" | Add-Content $logfile
       
        #Check for old CollectFiles Folder from previous run, remove if found. Check forms and instances folders exist after copy from device and perform ODK Briefcase export ready for Tariff processing
        
        #Lookup folder/directory name on local PC
        if((Test-Path $collectDir)){
            $localodkdirs = Get-ChildItem -Path $collectdir -Recurse -include forms,instances | Where-Object { $_.PSIsContainer }

            #Check if local Processing\CollectFiles Directory contains forms and instances and delete ready for processing
                ForEach ($localdir in $localodkdirs){
                    $localdirname = $collectDir + "\" + $localdir.Name
                    if((Test-Path $localdirname)){
                        Remove-Item $localdirname -Force -Recurse
                        "Directory $localdirName removed, ready for new entries from Device`r" | Add-Content $logfile
                    }
                }
        }
        $adbpull = Start-Process -FilePath adb.exe -ArgumentList "pull /sdcard/$odkdir/ $collectDir" -NoNewWindow -Wait -PassThru -RedirectStandardError $logfiledir\adblog.txt
        $adbpull
        if($adbpull.ExitCode -eq 1){
            ForEach ($line in (Get-Content $logfiledir\adblog.txt | Where-Object {$_ -like '*does not exist*'})) {
                $errors += $line
                "$errors : Unable to locate ODK Folder on the Android Device, Check ODKDir setting in config.txt matches the name of the ODK folder on the device`r" | Add-Content $logfile
                (new-object -ComObject wscript.shell).Popup("$errors `n`nUnable to locate ODK Folder on the Android Device, Check ODKDir setting in config.txt matches the name of the ODK folder on the device",0,"Error!")
                exit
            }
        }else{
            ForEach ($object in (Get-Content $logfiledir\adblog.txt | Where-Object {$_ -like '*files pulled.*'})) {
                $copyinfo += $object
                "$copyinfo `r" | Add-Content $logfile
                break
            }
        }

    }else{
        "Android Device not found. Check that it is connected, that Developer Options is enabled on the tablet (with USB Debugging mode turned on) and that the ADB drivers are up-to-date`r" | Add-Content $logfile
        (new-object -ComObject wscript.shell).Popup("Android Device not found. Check that it is connected, that Developer Options is enabled on the tablet (with USB Debugging mode turned on) and that the ADB drivers are up-to-date",0,"Error!")        
        exit
    }

}


#Lookup folder/directory name on local PC
$localodkdirs = Get-ChildItem -Path $collectDir -Recurse -include forms,instances | Where-Object { $_.PSIsContainer }

#Check if local Processing\CollectFiles Directory contains forms and instances folders after copy from device
            ForEach ($localdir in $localodkdirs){
                $localdirname = $collectDir + "\" + $localdir.Name
                if((Test-Path $localdirname)){
                    "Directory $localdirName confirmed available for conversion`r" | Add-Content $logfile
                }else{
                    $localpathfound = "0"
                    "Directory $localdirName not found, check that files have copied from device successfully`r" | Add-Content $logfile
                    (new-object -ComObject wscript.shell).Popup("Directory $localdirName not found, check that files have copied from device successfully",0,"Error!")
                    exit
                }
            }

if($localpathfound -ne 0){

    if(Test-Path $odkbcdir){
        Remove-Item $odkbcdir -Force -Recurse
        "ODK Briefcase Storage Directory cleared ready for processing`r" | Add-Content $logfile
    }
        
    Add-Type -assembly System.Windows.Forms

    ## -- Create The Progress-Bar
    $ObjForm = New-Object System.Windows.Forms.Form
    $ObjForm.Text = "SmartVA Processing Entries..."
    $ObjForm.Height = 100
    $ObjForm.Width = 500
    $ObjForm.BackColor = "white"
    $ObjForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $ObjForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    
    ## -- Create The Label
    $ObjLabel = New-Object System.Windows.Forms.Label
    $ObjLabel.Text = "Starting. Please wait ... "
    $ObjLabel.Left = 5
    $ObjLabel.Top = 10
    $ObjLabel.Width = 500 - 20
    $ObjLabel.Height = 15
    $ObjLabel.Font = "Tahoma"

    ## -- Add the label to the Form
    $ObjForm.Controls.Add($ObjLabel)
    $PB = New-Object System.Windows.Forms.ProgressBar
    $PB.Name = "PowerShellProgressBar"
    $PB.Value = 100
    $PB.Style="Continuous"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 500 - 40
    $System_Drawing_Size.Height = 20
    $PB.Size = $System_Drawing_Size
    $PB.Left = 5
    $PB.Top = 40
    $ObjForm.Controls.Add($PB)
    
    # -- Show the Progress-Bar and Start The PowerShell Script
    $ObjForm.Show() | Out-Null
    $ObjForm.Focus() | Out-NUll
    $ObjLabel.Text = "Starting. Please wait ... "
    $ObjForm.Refresh()
    Start-Sleep -milliseconds 150
    
    $ObjLabel.Text = "ODK Briefcase Conversion Started ... "
    $PB.Value = 20
    $ObjForm.Refresh()
    Start-Sleep -Milliseconds 150
        
    $bproc = Start-Process -FilePath java.exe -ArgumentList "-jar briefcase.jar -em -oc -od $collectDir -id $briefcaseFormID -sd $briefcaseStorageDir -ed `"$briefcaseExportDir`" -f $briefcaseConvertedFile" -NoNewWindow -Wait -PassThru -RedirectStandardError $logfiledir\BriefcaseOutput.txt
    $bproc
    "ODK Briefcase conversion Started...." | Add-Content $logfile

    if(!(Test-Path $odkbcdir)){
        "ODK Briefcase conversion is unable to run.  Please confirm Java is installed and that the system PATH environment variable is set.`r" | Add-Content $logfile
        (new-object -ComObject wscript.shell).Popup("ODK Briefcase conversion is unable to run.  Please confirm Java is installed and that the system PATH environment variable is set.",0,"Error!")
        $ObjForm.Close()
        $ObjForm.Dispose()
        exit
    }
    
    if($bproc.ExitCode -eq 1){
        ForEach ($line in (Get-Content $logfiledir\BriefcaseOutput.txt | Where-Object {$_ -like '*SEVERE*'})) {
            $errors += $line
            (new-object -ComObject wscript.shell).Popup("$errors `n`nCheck ODK Briefcase Storage locations match script expected location and that the FormID is correct",0,"Error!")
            "$errors : Check ODK Briefcase Storage locations match script expected location and that the FormID is correct`r" | Add-Content $logfile
            $ObjForm.Close()
            $ObjForm.Dispose()
            exit
        }
    }else{
        ForEach ($line in (Get-Content $logfiledir\BriefcaseOutput.txt | Where-Object {$_ -like '*SEVERE*'})) {
            $errors += $line
            (new-object -ComObject wscript.shell).Popup("$errors `n`nCheck config.txt file for correct FormID value",0,"Error!")
            "$errors : Check config.txt file for correct FormID value`r" | Add-Content $logfile
            $ObjForm.Close()
            $ObjForm.Dispose()
            exit
        }
        
        if($errors -eq $null){
            "Export File $briefcaseExportDir\$briefcaseConvertedFile created`r" | Add-Content $logfile
            "ODK Briefcase conversion completed...." | Add-Content $logfile
            
            #Clear Tariff Folder of old files for processing new tablet entries
            if(Test-Path $tariffOutput){
                Remove-Item $tariffOutput -Force -Recurse
                "Tariff Directory cleared ready for processing`r" | Add-Content $logfile
            }
            New-Item $tariffOutput -ItemType directory                
            
            #check Briefcase output file has any entries.  If not, advise the user and exit processing.
            if((Import-Csv $tariffInput) -eq $null){
                (new-object -ComObject wscript.shell).Popup("No data was found after the ODK Briefcase export.`n`nThis may be because there are no entries on the tablet for the specified form, or due to an export error.`n`nConfirm entries exist, that the FormID is correct in the config.txt file, or check the BriefcaseOutput.txt log for additional information.",0,"Error!")
                "No data was found after the ODK Briefcase export.`n`nThis may be because there are no entries on the tablet for the specified form, or due to an export error.`n`nConfirm entries exist, that the FormID is correct in the config.txt file, or check the BriefcaseOutput.txt log for additional information." | Add-Content $logfile
                $ObjForm.Close()
                $ObjForm.Dispose()
                exit
            }
            
            "Tariff Analysis started at $(Get-date)" | Add-Content $logfile

            $ObjLabel.Text = "Tariff Analysis Started ... "
            $PB.Value = 35
	        $ObjForm.Refresh()
            Start-Sleep -Milliseconds 150
            $tproc = Start-Process -FilePath SmartVA-Analyze-cli.exe -ArgumentList "--chinese --country $tariffCountry --hiv $tariffHIV --malaria $tariffMalaria --hce $tariffHCE --freetext $tariffFreetext --figures $tariffFigures `"$tariffInput`" $tariffOutput" -NoNewWindow -PassThru -RedirectStandardError $logfiledir\TariffOutput.txt 
            $tproc
            $ProcessList = "SmartVA-Analyze-cli"
            Do {
                $Percentage = $PB.Value
	            $ProcessesFound = Get-Process | ? {$ProcessList -contains $_.Name} | Select-Object -ExpandProperty Name
	                If ($ProcessesFound) {
                        $counter++
                        $ObjLabel.Text = "Tariff Analysis Processing ... $Percentage% "
                        [int]$PB.Value = 100 - ((50/($counter*.5)))
                        $ObjForm.Refresh()
                        if($counter -lt 6){
                            Start-Sleep 20
                        }elseif($counter -gt 5 -AND $counter -lt 21){
                            Start-Sleep 8
                        }else{
                            Start-Sleep 1
                        }
	                }
                $ObjForm.Refresh()
            } Until (!$ProcessesFound)
            $ObjForm.Close()
            $ObjForm.Dispose()
            
            #Check TariffOutput file for success or failure and store Archive copy of tablet files in Errors or Processed
            if(!(Get-Content $logfiledir\TariffOutput.txt | Where-Object {$_ -like '*Process Completed*'})){
                                
                $tariffError = Get-Content $logfiledir\TariffOutput.txt
                (new-object -ComObject wscript.shell).Popup("Tariff did not complete due to:`n`n$TariffError`n`n No Tariff Output is possible at this time and no items have been archived.`n`nPlease review error in $logfile and correct the issue before re-trying analysis.",0,"Error!")
                "$TariffError`r" | Add-Content $logfile
            
            }else{
                
                "Tariff analysis completed at $(Get-date)`r" | Add-Content $logfile
                #Update Archive with Forms/Instances that are new or have been updated since last Archive run.
                

                # get collectfiles form information (copy of Tablet form file)
	            $collectFormFile = Get-childitem $collectFormsDir -filter "$briefcaseFormID.xml"
	            
                # get archived form information (returns NULL if file is not found)
                $archiveFormFile = Get-childitem "$processedArchiveFormsDir" -filter "$briefcaseFormID.xml"

                #If no archive form file is found, copy form file to archive. If archive form file found, check tablet form file against existing archive file and update if a newer version is found on the tablet (to maintain Archive current version).  Alert user if tablet form file of the same name is older than existing archive file (may indicate only some users have updated the tablet form definitions to the latest version).
                if($archiveFormFile.Fullname -ne $null){
                    "Archive form file found, checking for newer version..." | Add-Content $logfile
                    
                    if(($collectFormFile.LastWriteTime.Date -gt $archiveFormFile.LastWriteTime.Date) -or ($collectFormFile.LastWriteTime.Date -eq $archiveFormFile.LastWriteTime.Date -and $collectFormFile.LastWriteTime.TimeofDay -gt $archiveFormFile.LastWriteTime.TimeofDay)){
                        
                        "Source Form file is newer than an existing Archive Form file of the same name...updating archive" | Add-Content $logfile
                        
                        try{
                            Copy-Item -Path $collectFormFile.FullName -Destination "$processedArchiveFormsDir" -Force -ErrorAction Stop
                            "Form File successfully Archived to $processedArchiveFormsDir" | Add-Content $logfile
                        }
                        catch{
                            (new-object -ComObject wscript.shell).Popup("An issue was encountered while copying to the Archive.  Please see $($logfile) for additional information",0,"Error!")                       
                            "Unable to complete Archive Form update due to: $_.Exception.  Correct this error and re-run analysis to ensure Archive data is updated BEFORE removing VA's from the Tablet." | Add-Content $logfile
                        }

                    }elseif(($collectFormFile.LastWriteTime.Date -eq $archiveFormFile.LastWriteTime.Date -and $collectFormFile.LastWriteTime.TimeofDay -lt $archiveFormFile.LastWriteTime.TimeofDay)){
                        "Source Form file is older than an existing Archive Form file.  No archive action required" | Add-Content $logfile
                     }else{
                        "Source Form file being used is identical to archive form file, no archive action required" | Add-Content $logfile
                    }
                }else{
	                    
                    "Archive Form file not found.  Copying form file from Tablet to Archive" | Add-Content $logfile
                        
                    try{
                        Copy-Item -Path $collectFormFile.FullName -Destination "$processedArchiveFormsDir" -Force -ErrorAction Stop
                        "Form File successfully Archived to $processedArchiveFormsDir" | Add-Content $logfile
                    }
                    catch{
                        (new-object -ComObject wscript.shell).Popup("An issue was encountered while copying to the Archive.  Please see $($logfile) for additional information",0,"Error!")                       
                        "Unable to complete Archive Form update due to: $_.Exception.  Correct this error and re-run analysis to ensure Archive data is updated BEFORE removing VA's from the Tablet." | Add-Content $logfile
                    }

                }


                # get collectfiles instances information (copy of Tablet instances files/folders)
	            $collectInstanceFiles = Get-childitem $collectInstancesDir -filter "$briefcaseFormID*" | Where-Object { $_.PSIsContainer }
	            
                # get archived instances information (returns NULL if files/folders are not found)
                $archiveInstanceFiles = Get-childitem "$processedArchiveInstancesDir" -filter "$briefcaseFormID*" | Where-Object { $_.PSIsContainer }
                
                #If no archive instances are found, copy instances to archive. If archive instances found, check collect instances against existing archive instances and update if a newer version is found (to maintain Archive current version).  Alert user if tablet instance of the same name is older than existing archive file (may indicate old data being reviewed).
                if($archiveInstanceFiles -ne $null){
                    "Archive instances found, check what needs to be copied..." | Add-Content $logfile

                    ForEach($collectInstanceFile in $collectInstanceFiles){
                        if(!(Get-childitem $processedArchiveInstancesDir -filter $collectInstanceFile | Where-Object { $_.PSIsContainer })){
                            try{
                                Copy-Item -Path $collectInstanceFile.FullName -Destination "$processedArchiveInstancesDir" -Recurse -Force -ErrorAction Stop
                                "Instance $collectInstanceFile successfully Archived to $processedArchiveInstancesDir" | Add-Content $logfile
                            }
                            catch{
                                "Unable to complete Archive Instance update of $collectInstanceFile due to: $_.Exception.  Correct this error and re-run analysis to ensure Archive data is updated BEFORE removing VA's from the Tablet." | Add-Content $logfile
                                $errorcount ++                       
                            }
                        }else{
                            $archiveInstanceFile = Get-childitem $processedArchiveInstancesDir -filter $collectInstanceFile | Where-Object { $_.PSIsContainer }
                            
                            if(($collectInstanceFile.LastWriteTime.Date -gt $archiveInstanceFile.LastWriteTime.Date) -or ($collectInstanceFile.LastWriteTime.Date -eq $archiveInstanceFile.LastWriteTime.Date -and $collectInstanceFile.LastWriteTime.TimeofDay -gt $archiveInstanceFile.LastWriteTime.TimeofDay)){
                        
                                "Source Instance $collectInstanceFile is newer than an existing Archive Instance file of the same name...updating archive to reflect changes" | Add-Content $logfile
                        
                                try{
                                    Copy-Item -Path $collectInstanceFile.FullName -Destination "$processedArchiveInstancesDir" -Recurse -Force -ErrorAction Stop
                                    "Instance $collectInstanceFile successfully Archived to $processedArchiveInstancesDir" | Add-Content $logfile
                                }
                                catch{
                                    "Unable to complete Archive Instance update of $collectInstanceFile due to: $_.Exception.  Correct this error and re-run analysis to ensure Archive data is updated BEFORE removing VA's from the Tablet." | Add-Content $logfile
                                    $errorcount ++
                                }

                            }elseif(($collectInstanceFile.LastWriteTime.Date -eq $archiveInstanceFile.LastWriteTime.Date -and $collectInstanceFile.LastWriteTime.TimeOfDay -lt $archiveInstanceFile.LastWriteTime.TimeOfDay)){
                                "Source Instance $collectInstanceFile is older than an existing Archive Instance file.  Please check your dataset to ensure you have the latest copy of all data." | Add-Content $logfile
                                $errorcount ++
                            }else{
                                "Source Instance File $collectInstanceFile is identical to existing archive instance file, no archive action required" | Add-Content $logfile
                            }
                        }
                    }
                    if($errorcount -gt 0){
                        "Number of Files in Error: $errorcount" | Add-Content $logfile
                        (new-object -ComObject wscript.shell).Popup("An issue was encountered while copying to the Archive.`n`nReview and correct the error then re-run analysis to ensure Archive data is updated BEFORE removing VA's from the Tablet.`n`nFor error details see $($logfile).",0,"Error!")
                    }
                }else{
                    try{
                        ForEach($collectInstanceFile in $collectInstanceFiles){
                            Copy-Item -Path $collectInstanceFile.Fullname -Destination "$processedArchiveInstancesDir" -Recurse -Force -ErrorAction Stop
                            "Instances $($collectInstanceFile.Fullname) successfully Archived to $processedArchiveInstancesDir" | Add-Content $logfile
                        }
                    }
                    catch{
                        (new-object -ComObject wscript.shell).Popup("An issue was encountered while copying to the Archive.  Please see $($logfile) for additional information",0,"Error!")                       
                        "Unable to complete Archive Instances update due to: $_.Exception.  Correct this error and re-run analysis to ensure Archive data is updated BEFORE removing VA's from the Tablet." | Add-Content $logfile
                    }
                }
                #Lookup file names for all Tariff Predictions and current archive results. Display Tariff Predictions to screen and add to archive results.csv file if not already exists
                                
                $tariffpredictions = get-childitem "$tariffOutput" -Recurse | where {$_.name -like '*likelihoods.csv'}
                $archiveResults = Import-Csv "$resultsfile"

                if($tariffpredictions -ne $null){

                    ForEach ($prediction in $tariffpredictions){
                        $path = $prediction.FullName
                        $results = $null
                        $VAs = Import-Csv "$path"

                        ForEach ($VA in $VAs){
                            $x = 0
                            $sid = $($VA.sid)
                            $age = $($VA.age)
                            $sex = $($VA.sex)
                            if($sex -eq 2){ 
                                $gender = "Female"
                            }elseif($sex -eq 1){
                                $gender = "Male"
                            }elseif($sex -eq 3){
                                $gender = "Third Gender"
                            }elseif($sex -eq 8){
                                $gender = "Refused to answer"
                            }elseif($sex -eq 9){
                                $gender = "Don't know"
                            }else{
                                $gender = $sex
                            }
                            $cod1 = $($VA.cause1)
                            $lh1 = $($VA.likelihood1)
                            
                            $cod2 = $($VA.cause2)
                            $lh2 = $($VA.likelihood2)
                            
                            $cod3 = $($VA.cause3)
                            $lh3 = $($VA.likelihood3)
                            
                            $as = $($VA.all_symptoms)
                            $as_repl = $as -replace ";", "`r"
                            
                            
                            
                            $results += "VA ID:`t`t" + $sid + "`n年龄:`t`t" + $age + "`n性别:`t`t" + $gender + "`n`n死亡原因1:`t" + $cod1 + "`n可能性:`t" + $lh1 + "`n`n死亡原因2:`t" + $cod2 + "`n可能性:`t" + $lh2 + "`n`n死亡原因3:`t" + $cod3 + "`n可能性:`t" + $lh3 + "`n`n" + "全部症状:`n" + $as_repl
                            $sid + "," + $age + "," + $gender + "," + $cod1 + "," + $lh1 + "," + $cod2 + "," + $lh2 + "," + $cod3 + "," + $lh3 + "," + $as_repl | Add-Content $gridfile
                        
                            #add to results file
                            
                            if($archiveResults -ne $null){
                                ForEach ($archiveResult in $archiveResults){
                                    if ($sid -eq $archiveResult.VA_ID ){
                                        $x++
                                    }
                                }
                                if($x -eq 0){
                                    try{
                                        $sid + "," + $age + "," + $gender + "," + $cod1 + "," + $lh1 + "," + $cod2 + "," + $lh2 + "," + $cod3 + "," + $lh3 + "," + $datetime + "," | Add-Content $resultsfile -ErrorAction Stop
                                        "Result with VA_ID $sid successfully Archived to $resultsfile" | Add-Content $logfile
                                    }
                                    catch{
                                        $archiveResultErrorCount++
                                        $archiveErrorMsg = $null
                                        $archiveErrorMsg = $_
                                    }
                                }
                            }else{
                                    try{
                                        $sid + "," + $age + "," + $gender + "," + $cod1 + "," + $lh1 + "," + $cod2 + "," + $lh2 + "," + $cod3 + "," + $lh3 + "," + $datetime + "," | Add-Content $resultsfile -ErrorAction Stop
                                        "Result with VA_ID $sid successfully Archived to $resultsfile" | Add-Content $logfile
                                    }
                                    catch{
                                        $archiveResultErrorCount++
                                        $archiveErrorMsg = $null
                                        $archiveErrorMsg = $_
                                    }
                            }
                            $resultcount ++
                        }
                        if($archiveResultErrorCount -gt 0){
                            (new-object -ComObject wscript.shell).Popup("Could not write results to Archive file $resultsfile due to: `n`n$archiveErrorMsg`n`n If this file is open, please close it and repeat analysis to capture results in the archive",0,"Error")
                            "Could not write results to Archive file $resultsfile due to: $archiveErrorMsg" | Add-content $logfile
                            exit
                        }
                        $allresults += $results
                    }
                    if($resultcount -eq 1){
                        #Display Tariff Results to screen in easy to read format
                        (new-object -ComObject wscript.shell).Popup("$allresults",0,"Traslation Results")
                        #Ask for CoD as written on Death Cert
                        if($x -eq 0 -and $cod1 -ne "Undetermined"){
                            do{
                                $PhdUCoD = [Microsoft.VisualBasic.Interaction]::InputBox("请输入死亡证明书上填写的根本死因", "请输入死亡证明书上填写的根本死因", "")
                            }until($PhdUCoD -ne "")
                            "$sid" + "," + "$CoD1" + "," + "$PhdUCod"| Add-content $mccod
                        }
                    }else{
                        Import-Csv -Path $gridfile -Header "VA ID", Age, Sex, "Cause 1", "Likelihood 1", "Cause 2", "Likelihood 2", "Cause 3", "Likelihood 3", "All Symptoms" | Out-GridView -wait –Title "Results"
                    }
                }else{
                    (new-object -ComObject wscript.shell).Popup("No results to display.  This is normal if processed VA's had no age information, or consent was not provided.",0,"SmartVA Results")
                }
            }
        } 
    }
}