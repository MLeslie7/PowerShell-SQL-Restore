[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo") 
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SMO")
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SmoExtended") 

##########Constants######

[int]$TotalBackupFiles = 125 #Limit of rows to retrive for recovery, should slightly exceed # of backups taken for the required retention period

[string]$DefaultMailSQLServer = "[DB Mail Server]" # SQL Server with DB Mail setup and profile created

[string]$DefaultMailProfile = "[DB Mail Profile]" # SQL Server DB Mail profile setup for sending email

<#

$TotalBAckup Files is return limit in query to retrieve backup files, 
should be set by retention policy and frequency of backups
Example:

Full backup 1/week on Sunday
Differential 1/day except Sunday
Transaction log backups every 2 hours
1 week retention
1 + 6 + (7x12) = 91


Full backup 1/week on Sunday
Differential 1/day except Sunday
Transaction log backups every 1 hours
2 week retention
2 + 12 + (14x24) = 350

#>

##########End Constants##

$Error.Clear()
function CheckPSVersion{

    try{

        Write-Output "`nChecking PowerShell version...`n" 

        $MajorV = $PSVersionTable.PSVersion.Major
        $MinorV = $PSVersionTable.PSVersion.Minor

        if($MajorV -lt 4) {

            Write-Warning "PowerShell $MajorV.$MinorV detected, 4.0 or higher required"
            Write-Output "Install PowerShell 4.0 and then try again." 
            Read-Host "Press enter to continue"
            break;
        }

        if($error[0].Exception){Throw $error[0].Exception}
        
        Write-Output "PowerShell $MajorV.$MinorV Installed!`n"

    }catch{

        Write-Output "Error detecting PowerShell version update then try again."  
        break;

    }

}
CheckPSVersion

Function ShowPromptServerName{
   
    Param([Parameter(Mandatory=$true)][string]$PromptText)

    [string]$Server = Read-Host "`n$PromptText"
 
    return $Server
}
Function ConnectSQLServer{

    Param([Parameter(Mandatory=$true)][string]$SQLServer)

    #Connect to server
    $SQLCon = New-SMOconnection -server $SQLServer

    Return $SQLCon

}
Function New-SMOconnection {

    Param (
        [Parameter(Mandatory=$true)]
        [string]$server,
        [int]$StatementTimeout=0
    )

        $conn = New-Object Microsoft.SqlServer.Management.Common.ServerConnection($server)

        $conn.applicationName = "PowerShell SMO"

        $conn.StatementTimeout = $StatementTimeout

        Try {
    
            $conn.Connect()
            $smo = New-Object Microsoft.SqlServer.Management.Smo.Server($conn)
            $smo
        }

        Catch{

            if ($conn.IsOpen -eq $false) {

                Write-Warning "Could not connect to the SQL Server Instance $server. Try ServerName\InstanceName"
                $smo = "NOConnection"
                $smo        
            }
        }
  
}
Function Get-DatabaseList {
    
    Param (
        [Parameter(Mandatory=$true)]
        [Microsoft.SqlServer.Management.Smo.SqlSmoObject]$ServerConnection
    )

        $qry = "EXEC sp_databases"

        $db = $ServerConnection.Databases["master"]

        $rs = $db.ExecuteWithResults($qry)

        return $rs.Tables.Rows

}
Function Get-BackupFileList {

    Param(
        [Parameter(Mandatory=$true)][string]$DatabaseName,
        [Microsoft.SqlServer.Management.Smo.SqlSmoObject]$ServerConnection 
    )

    $query = @"
SELECT Top ($TotalBackupFiles) Name,backup_finish_date,type,database_name,server_name,physical_device_name,checkpoint_lsn,database_backup_lsn
FROM    msdb.dbo.backupset AS s WITH (nolock) INNER JOIN
            msdb.dbo.backupmediafamily AS f WITH (nolock) ON s.media_set_id = f.media_set_id
WHERE   (s.database_name = '$DatabaseName') AND (f.device_type <> 7) 
Order By s.backup_finish_date Desc
"@

        $db = $ServerConnection.Databases["msdb"]
    
        $rs = $db.ExecuteWithResults($query)

        return $rs.Tables.Rows


}
Function Get-ListToRestore {

    Param(
        [Parameter(Mandatory=$true)][string]$DatabaseName,
        [Parameter(Mandatory=$true)][datetime]$DatabaseBackupFinish,
        [Parameter(Mandatory=$true)][decimal]$Backuplsn,
        [Microsoft.SqlServer.Management.Smo.SqlSmoObject]$ServerConnection
    )
     
   
$query = @"
DECLARE @dbname varchar(100), @DBFinish datetime,@Backuplsn numeric (38,0), @LastIncrmnt datetime, @SelectedType char(1)
Declare  @FileTable Table(Name nvarchar(128),checkpoint_lsn numeric(38,0),
database_backup_lsn numeric(38,0),backup_finish_date datetime ,type char(1)
,database_name nvarchar(128),server_name nvarchar(128),physical_device_name nvarchar(260))
SET @dbname = '$DatabaseName'
set @DBFinish = '$DatabaseBackupFinish'
Set @Backuplsn = $Backuplsn

Select @SelectedType = type
FROM    msdb.dbo.backupset AS s WITH (nolock) INNER JOIN
            msdb.dbo.backupmediafamily AS f WITH (nolock) ON s.media_set_id = f.media_set_id
WHERE   (s.database_name = @dbname) AND (s.backup_finish_date = @DBFinish) 
		AND ((s.database_backup_lsn = @Backuplsn) or (s.checkpoint_lsn = @Backuplsn))
		 AND (f.device_type <> 7) AND (s.is_snapshot = 0)

Insert Into @FileTable
SELECT Top(1) Name,checkpoint_lsn,database_backup_lsn,backup_finish_date,type,database_name,server_name,physical_device_name
FROM    msdb.dbo.backupset AS s WITH (nolock) INNER JOIN
            msdb.dbo.backupmediafamily AS f WITH (nolock) ON s.media_set_id = f.media_set_id
WHERE   (s.database_name = @dbname) AND (s.backup_finish_date <= @DBFinish) 
		AND ((s.database_backup_lsn = @Backuplsn) or (s.checkpoint_lsn = @Backuplsn)) AND (s.type = 'd')
		 AND (f.device_type <> 7) AND (s.is_snapshot = 0)
Order By s.backup_finish_date Desc

if( @SelectedType <> 'D') Begin --if not type full or 'd'

Insert Into @FileTable
SELECT Top(1) Name,checkpoint_lsn,database_backup_lsn,backup_finish_date,type,database_name,server_name,physical_device_name
FROM    msdb.dbo.backupset AS s WITH (nolock) INNER JOIN
            msdb.dbo.backupmediafamily AS f WITH (nolock) ON s.media_set_id = f.media_set_id
WHERE   (s.database_name = @dbname) AND (s.backup_finish_date <= @DBFinish) 
		AND (s.database_backup_lsn = @Backuplsn)  AND (s.type = 'i')
		 AND (f.device_type <> 7) AND (s.is_snapshot = 0)
Order By s.backup_finish_date Desc

Select @LastIncrmnt = backup_finish_date
FROM    msdb.dbo.backupset AS s WITH (nolock) INNER JOIN
            msdb.dbo.backupmediafamily AS f WITH (nolock) ON s.media_set_id = f.media_set_id
WHERE   (s.database_name = @dbname) AND (s.backup_finish_date <= @DBFinish) 
		AND ((s.database_backup_lsn = @Backuplsn)or (s.checkpoint_lsn = @Backuplsn)) AND (s.type <> 'l')
		 AND (f.device_type <> 7) AND (s.is_snapshot = 0)
Order By s.backup_finish_date Asc

Insert Into @FileTable
SELECT  Name,checkpoint_lsn,database_backup_lsn,backup_finish_date,type,database_name,server_name,physical_device_name 
FROM    msdb.dbo.backupset AS s WITH (nolock) INNER JOIN
            msdb.dbo.backupmediafamily AS f WITH (nolock) ON s.media_set_id = f.media_set_id
WHERE   (s.database_name = @dbname) AND (@DBFinish >= s.backup_finish_date) AND (s.backup_finish_date > @LastIncrmnt)
		AND (s.database_backup_lsn = @Backuplsn)  AND (s.type = 'l')
		 AND (f.device_type <> 7) AND (s.is_snapshot = 0)
Order By s.backup_finish_date Asc

end --if not type full or 'd'

Select * From @FileTable
"@

    $db = $ServerConnection.Databases["msdb"]
    
    $rs = $db.ExecuteWithResults($query)

    return $rs.Tables.Rows

}
function Get-RestoreFolderLocation {

    Param (

        [Parameter(Mandatory=$true)][string]$DisplayPrompt,
        [Parameter(Mandatory=$true)][string]$PathSelected

    )
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = $DisplayPrompt
    $foldername.rootfolder = "MyComputer"
    $foldername.SelectedPath = $PathSelected

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }

    return $folder

}
Function Restore-FilesJob{

    Param(

    [Parameter(Mandatory=$true)][psobject]$FileList,
    [Parameter(Mandatory=$true)][string]$RestoreName

    )

    if(!$FileList.count){

        $RecovOpt = "RECOVERY";
        $Filecount = 1;
        $CountFiles = $false;    

    }else{

        $RecovOpt="NORECOVERY";
        $Filecount = $FileList.Count;
        $CountFiles = $true;

    }

    $stTime = (Get-Date -Format "M-dd-yyyy HH:mm")

    Write-Warning "Wait for job to complete..."
    Write-Output "Job started at $stTime"

    $Error.Clear()

    ForEach($BFile in $FileList){

        if($Error){$SourceSQLCon.ConnectionContext.Disconnect();$DestSQLCon.ConnectionContext.Disconnect();break;}

        if($CountFiles -eq $true){

            $PlacePos = ($FileList.IndexOf($BFile)+1)
            if($PlacePos -eq $Filecount){$RecovOpt = "RECOVERY"}

        }else{

            $PlacePos = 1
        }

        if($RecoveryOption -eq "NoRecovery"){$RecovOpt = "NORECOVERY"}

        switch($BFile.type){

            "l"{
                Write-Output "`nRestoring Transaction Log backup file [$PlacePos] of [$Filecount] in job:"$BFile.physical_device_name; 
                Restore-TLogBackup -FileName $Bfile.physical_device_name -DBRestoreName $RestoreName `
                    -RecoverySw $RecovOpt -ServerConnection $DestSQLCon 
                }

            "i"{
                Write-Output "`nRestoring Differential backup file [$PlacePos] of [$Filecount] in job:"$BFile.physical_device_name;
                Restore-FullBackup -FileName $Bfile.physical_device_name -DBCurrentName $BFile.database_name `
                    -DBRestoreName $RestoreName -MDFPath $RestoreMDFFolder -ldfPath $RestoreLDFFolder `
                    -MDFLogicalName $LogicalNames.LogicalDataFile -ldfLogicalName $LogicalNames.LogicalLogFile `
                    -RecoverySw $RecovOpt -ServerConnection $DestSQLCon
                }

            "d"{
                Write-Output "`nRestoring Full backup file [$PlacePos] of [$Filecount] in job:"$BFile.physical_device_name;
                Restore-FullBackup -FileName $Bfile.physical_device_name -DBCurrentName $BFile.database_name `
                    -DBRestoreName $RestoreName -MDFPath $RestoreMDFFolder -ldfPath $RestoreLDFFolder `
                    -MDFLogicalName $LogicalNames.LogicalDataFile -ldfLogicalName $LogicalNames.LogicalLogFile `
                    -RecoverySw $RecovOpt -ServerConnection $DestSQLCon
                } 
        }
        
    }

    $finTime = (Get-Date -Format "M-dd-yyyy HH:mm")
    

    if(!($Error)){

        Write-Output "`nJob has completed at $finTime check the server $DestSQLsvr for the database $RestoreName"

        if($SendEmail -eq $true){

            $DBMailData = "<h2>Database Restore: $RestoreName</h2><h3>To Server: $DestSQLsvr Completed Successfully</h3><h3>Start Time: $stTime</br>Finish Time: $finTime</h3>"
            SendDBMail -DBMailServer $DBMailSrvr -DBMailProfile $DBMailProfile -EmailRecipients $DBMailRecipients -EmailTableData $DBMailData
        }
        
    }else{

        Write-Warning "`Error during the restore finished at $finTime, verify input and try again."

        if($SendEmail -eq $true){

            $DBMailData = "<h2>Database Restore: $RestoreName</h2<h3>Error Encountered during the Restore</h3><h3>Start Time: $stTime</br>Finish Time: $finTime</h3>"
            SendDBMail -DBMailServer $DBMailSrvr -DBMailProfile $DBMailProfile -EmailRecipients $DBMailRecipients -EmailTableData $DBMailData
        }
    }

}
Function Restore-FullBackup{

    Param(

        [Parameter(Mandatory=$true)][string]$FileName,
        [Parameter(Mandatory=$true)][string]$DBCurrentName,
        [Parameter(Mandatory=$true)][string]$DBRestoreName,
        [Parameter(Mandatory=$true)][string]$MDFPath,
        [Parameter(Mandatory=$true)][string]$ldfPath,
        [Parameter(Mandatory=$true)][string]$MDFLogicalName,
        [Parameter(Mandatory=$true)][string]$ldfLogicalName,
        [Parameter(Mandatory=$true)]
        [ValidateSet("RECOVERY","NORECOVERY")]
        [string]$RecoverySw,
        [Microsoft.SqlServer.Management.Smo.SqlSmoObject]$ServerConnection 

    )

    $Error.Clear()

    $query = @"
RESTORE DATABASE $DBRestoreName
FROM disk='$FileName'
WITH $RecoverySw,
MOVE '$MDFLogicalName' TO 
'$MDFPath\$DBRestoreName.mdf', 
MOVE '$ldfLogicalName' 
TO '$ldfPath\$DBRestoreName.ldf'
"@

    $db = $ServerConnection.Databases["master"]
    
    $rs = $db.ExecuteWithResults($query)

    if($Error){

        Write-Warning "Error executing query`n$query."
        
    }else{

        Write-Output "`nRestored file $MDFPath\$DBRestoreName.mdf with $RecoverySw"
        Write-Output "Restored file $ldfPath\$DBRestoreName.ldf with $RecoverySw"

    }

    return $rs.Tables.Rows
    
}
Function Restore-TLogBackup {
    
    Param(    
        [Parameter(Mandatory=$true)][string]$FileName,
        [Parameter(Mandatory=$true)][string]$DBRestoreName,
        [Parameter(Mandatory=$true)]
        [ValidateSet("RECOVERY","NORECOVERY")]
        [string]$RecoverySw,
        [Microsoft.SqlServer.Management.Smo.SqlSmoObject]$ServerConnection 

    )

    $Error.Clear()

    $query = @"
RESTORE LOG $DBRestoreName
FROM disk='$FileName'
WITH $RecoverySw
"@

    $db = $ServerConnection.Databases["master"]
    
    $rs = $db.ExecuteWithResults($query)

    if($Error){

        Write-Warning "Error executing query`n$query."

    }else{

        Write-Output "`nRestored transaction log file for database $DBRestoreName with $RecoverySw"

    }

    return $rs.Tables.Rows
    #Write-Output $query

}
Function Get-LogicalFileNames {

    Param (

        [Parameter(Mandatory=$true)][string]$DBName,
        [Microsoft.SqlServer.Management.Smo.SqlSmoObject]$ServerConnection
    )
    
    
    $query = @"
Declare @DBName sysname,@LogicalDataFile sysname,@LogicalLogFile sysname,
        @PhysicalDataFile nvarchar(260),@PhysicalLogFile nvarchar(260)
 
set @DBName = '$DBName'
 
select  @LogicalDataFile = name, @PhysicalDataFile = physical_name
from    sys.master_files
where   database_id = db_id(@DBName) and type_desc = 'ROWS'
 
select  @LogicalLogFile = name, @PhysicalLogFile = physical_name
from    sys.master_files
where   database_id = db_id(@DBName) and type_desc = 'LOG'
 
select  @LogicalDataFile as [LogicalDataFile], @LogicalLogFile as [LogicalLogFile],
        @PhysicalDataFile as [PhysicalDataFile], @PhysicalLogFile as [PhysicalLogFile]
"@

    $db = $ServerConnection.Databases["master"]
    
    $rs = $db.ExecuteWithResults($query)

    return $rs.Tables.Rows


}
Function PromptForEmailProfile {

    $title = "Send Email on job completion"
    $message = "Send email with job results when job is completed showing errors if any.
Database mail must be set-up for this feature to function properly."

    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes send email", `
        "Enter yes to enter profile name and recipients of completion email."

    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No do not send email", `
        "Job will run but no email will be sent."

    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

    $result = $host.ui.PromptForChoice($title, $message, $options, 0) 

    switch ($result){

        0{Return $true}
        1{Return $false}

    }

}
Function PromptRecoveryOption {

    $title = "Recovery Option for restore"
    $message = "Set RECOVERY - NORECOVERY option for restore, Recovery is default and will allow processing requests when completed.
NoRecovery will leave the database in recovering mode and not able to process requests.
Use default - Recovery if not sure."

    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Recovery", `
        "Process backup files and the database will be left in Recovered state able to processes transactions."

    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&NoRecovery", `
        "Process backup files and leave the database in Recovering mode, unable to processes transactions."

    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

    $result = $host.ui.PromptForChoice($title, $message, $options, 0) 

    switch ($result){

        0{Return "Recovery"}
        1{Return "NoRecovery"}

    }

}
Function SendDBMail{

     Param (
        [Parameter(Mandatory=$true)][string]$DBMailServer,
        [Parameter(Mandatory=$true)][string]$DBMailProfile,
        [Parameter(Mandatory=$true)][string]$EmailRecipients,
        [Parameter(Mandatory=$true)][string]$EmailTableData
    )

    $SQLCon = ConnectSQLServer -SQLServer $DBMailServer

$query =@"
EXEC msdb.dbo.sp_send_dbmail 
@profile_name = '$DBMailProfile',
@recipients='$EmailRecipients',
@subject = 'Restore Job Results',
@body_format = 'HTML',
@body = '$EmailTableData';
"@       


    $db = $SQLCon.Databases["master"]

    $rs = $db.ExecuteWithResults($query)

    $SQLCon.ConnectionContext.Disconnect()

    return $rs.Tables.Rows

}
Function Get-MailProfiles{
    
    Param (
        [Parameter(Mandatory=$true)]
        [Microsoft.SqlServer.Management.Smo.SqlSmoObject]$ServerConnection
    )

        $qry = "EXECUTE msdb.dbo.sysmail_help_profileaccount_sp"

        $db = $ServerConnection.Databases["msdb"]

        $rs = $db.ExecuteWithResults($qry)

        return $rs.Tables.Rows
}
Function PromptForConfirmation {
    
    #get logical file names for restore command
    Write-Output "`nGetting logical file names and paths of existing database..."

    $LogicalNames = Get-LogicalFileNames -DBName $selectedBackup.database_name -ServerConnection $SourceSQLCon

    $LogicalNames | ft -AutoSize -Wrap

    #display list of files for confirmation
    $FilestoRestoreList | Select-Object database_name,type,backup_finish_date,physical_device_name,server_name | ft -AutoSize -Wrap

    $dbname = $selectedBackup.database_name
    $recvryTime = $selectedBackup.backup_finish_date.ToString("yyyy-MM-dd HH:mm:ss")

    Write-Output "Source server:         $SourceServer"
    Write-Output "Destination server:    $DestSQLsvr"
    Write-Output "MDF file location:     $RestoreMDFFolder"
    Write-Output "LDF file location:     $RestoreLDFFolder"
    Write-Output "Source database:       $dbname"
    Write-Output "Restore name:          $RestoreDBName"
    Write-Output "Recovery time:         $recvryTime"
    Write-Output "Recovery mode:         $RecoveryOption"

    if($SendEmail -eq $true){

        Write-Output "DB mail server:        $DBMailSrvr"
        Write-Output "DB mail profile:       $DBMailProfile"
        Write-Output "DB mail recipients:    $DBMailRecipients"

    }else{

        Write-Output "Send DB mail:          False"
    }
    
    $title = "Restore Database?"
    $message = "Confirm the details above before beginning the restore"

    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
        "Restores files listed to the location and names the database the Restore name."

    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
        "Cancels and ends script."

    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

    $result = $host.ui.PromptForChoice($title, $message, $options, 0) 

    switch ($result)
        {
            0 {Restore-FilesJob -FileList $FilestoRestoreList -RestoreName $RestoreDBName}
            1 {$SourceSQLCon.ConnectionContext.Disconnect();$DestSQLCon.ConnectionContext.Disconnect();break;}
        }

}

$Error.Clear()

do{

    $Error.Clear()

    $SourceServer = ShowPromptServerName -PromptText "Enter name of source SQL Server that backups were created on"

    Write-Output "`nAttempting SQL Connection to $SourceServer, please wait..."

    $SourceSQLCon = ConnectSQLServer -SQLServer $SourceServer

}until($SourceSQLCon -ne "NOConnection")
if($Error){

    Write-Warning "An error has occurred, try again."
    $SourceSQLCon.ConnectionContext.Disconnect();break;

}else{

    #Get list of Databases and display prompt to select database
    $SrcDataBase = Get-DatabaseList -ServerConnection $SourceSQLCon | Out-GridView -Title "Select Database to view backup files" -OutputMode Single

}
if($Error){

    Write-Warning "An error has occurred, try again."
    $SourceSQLCon.ConnectionContext.Disconnect();break;

}else{

    #Get list of backup files on selected server for selected database
    $BuFiles = Get-BackupFileList -DatabaseName $SrcDataBase.DATABASE_NAME -ServerConnection $SourceSQLCon
}
if($Error){

    Write-Warning "An error has occurred, try again."
    $SourceSQLCon.ConnectionContext.Disconnect();break;

}else{

    #Display backup files and prompt for recovery point
    $selectedBackup = $BuFiles | Out-GridView -Title "Select Backup closest to Recover Point Time" -OutputMode Single
}
if($Error){

    Write-Warning "An error has occurred, try again."
    $SourceSQLCon.ConnectionContext.Disconnect();break;

}else{

    #Query for files to restore
    $FilestoRestoreList = Get-ListToRestore -DatabaseName $selectedBackup.database_name `
                            -DatabaseBackupFinish $selectedBackup.backup_finish_date.ToString("yyyy-MM-dd HH:mm:ss") `
                            -Backuplsn $selectedBackup.database_backup_lsn `
                            -ServerConnection $SourceSQLCon
}
if($Error){

    Write-Warning "An error has occurred, try again."
    $SourceSQLCon.ConnectionContext.Disconnect();break;

}else{

    #Display files to restore
    $FilestoRestoreList | Out-GridView -Title "The listed files will be restored in order close window to continue" -Wait
}
if($Error){

    Write-Warning "An error has occurred, try again."
    $SourceSQLCon.ConnectionContext.Disconnect();break;

}else{

    do{
    
        $Error.Clear()

        #prompt for destination SQL Server
        $DestSQLsvr = ShowPromptServerName -PromptText "Enter the name of the Destination SQL server, the server the database will be restored to" #"lab-manager\veeamsql2012" #
    
        Write-Output "`nAttempting SQL Connection to $DestSQLsvr, please wait..."

        #connect to to dest server
        $DestSQLCon = ConnectSQLServer -SQLServer $DestSQLsvr

    }until($DestSQLCon -ne "NoConnection")

}
if($Error){

    Write-Warning "An error has occurred, try again."
    $SourceSQLCon.ConnectionContext.Disconnect();$DestSQLCon.ConnectionContext.Disconnect();break;

}else{

    Do{

    $Error.Clear()

    $FolderSel = $true

    #if destination SQL Server is this machine show choice to allow using folder dialog or entering network path
    if($DestSQLsvr.Split('\')[0] -eq $env:COMPUTERNAME){

    $title = "Restore Database to a network share or local folder?"
    $message = "`nChoose where to restore the database files to by entering your choice, 
you cannot restore files to the root of a drive it must be a folder.`n"

    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Local Folder", `
        "Use a folder browser to select location to save files to."

    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&Network Share", `
        "Enter a network share path to save files to."

    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

    $result = $host.ui.PromptForChoice($title, $message, $options, 0) 

    switch ($result)
        {
            #if Local folder selected then show folder dialog
            0 {
            
                #Prompt for Restore mdf Destination folder
                $RestoreMFolder = Get-RestoreFolderLocation -DisplayPrompt "Select Folder to save MDF file to" -PathSelected "C:\"
                $RestoreMDFFolder = $RestoreMFolder[1] #"c:\sql" #

                #Prompt for Restore mdf Destination folder
                $RestoreLFolder = Get-RestoreFolderLocation -DisplayPrompt "Select Folder to save Log file to" -PathSelected $RestoreMDFFolder
                $RestoreLDFFolder = $RestoreLFolder[1] #"C:\sql\log" #
            
            }

            #if network path selected prompt for location
            1 {
            
                #if restoring to remote machine prompt to manually enter path
                $RestoreMDFFolder = Read-Host "`nEnter restore path for MDF on $DestSQLsvr (\\server\share)"
                $RestoreLDFFolder = Read-Host "`nEnter restore path for LDF on $DestSQLsvr (\\server\share)"
            
            }
        }     

    }else{
        
        #if restoring to remote machine prompt to manually enter path
        $RestoreMDFFolder = Read-Host "`nEnter restore path for MDF on $DestSQLsvr (C:\RestoreFolder or \\server\share)"
        $RestoreLDFFolder = Read-Host "`nEnter restore path for LDF on $DestSQLsvr (C:\RestoreFolder or \\server\share)"

    }

    #remove trailing '\' if entered manually or root of drive selected
    $RestoreMDFFolder = $RestoreMDFFolder.TrimEnd("\")
    $RestoreLDFFolder = $RestoreLDFFolder.TrimEnd("\")

    #if root of a drive selected throw error
    If($RestoreMDFFolder.Length -le 2) {

        Write-Warning "You cannot restore the data file to the root of a drive, select or enter a folder to continue."
        $FolderSel = $false
    }

    #if root of a drive selected throw error
    If($RestoreLDFFolder.Length -le 2){

         Write-Warning "You cannot restore the log file to the root of a drive, select or enter a folder to continue."
         $FolderSel = $false
    }

}until($FolderSel -eq $true)

}
if($Error){

    Write-Warning "An error has occurred, try again."
    $SourceSQLCon.ConnectionContext.Disconnect();$DestSQLCon.ConnectionContext.Disconnect();break;

}else{

    $RecoveryOption = PromptRecoveryOption

}
if($Error){

    Write-Warning "An error has occurred, try again."
    $SourceSQLCon.ConnectionContext.Disconnect();$DestSQLCon.ConnectionContext.Disconnect();break;

}else{

    $title = "Append Restored Database Name?"
    $message = "`nChose a custom string to append to the end of the database name,
the date-time the restore starts, the restore point time, or nothing appended to the database
name. The maximum length for the Name is 128 characters, do not use special characters in a custom
string. If you choose to not append anything to the name, be sure the database does not exist on
the destination server, it will not be over written.`n"

    $c1 = New-Object System.Management.Automation.Host.ChoiceDescription "&Date-Time restore starts", `
        "The Date-Time the restore job begins will be appended to the database name."

    $c2 = New-Object System.Management.Automation.Host.ChoiceDescription "&Custom string", `
        "Enter a custom string such as 'Restored' to the database name."

    $c3 = New-Object System.Management.Automation.Host.ChoiceDescription "&Restore time", `
        "The time of the latest backup will be appended to the end of the database name."

    $c4 = New-Object System.Management.Automation.Host.ChoiceDescription "&Use existing name", `
        "Nothing will be appended to the end of the database name Be sure the database does not exist on the destination server."

    $options = [System.Management.Automation.Host.ChoiceDescription[]]($c1,$c2,$c3,$c4)

    $result = $host.ui.PromptForChoice($title, $message, $options, 0)

    switch($result){


        0{$AddStr = (Get-Date -Format MddyyyyHHmm)}
        1{$AddStr = Read-Host "Enter a string to append to the end of the database name" }
        2{$AddStr = $selectedBackup.backup_finish_date.ToString("MMddyyyyHHmm")}
        3{$AddStr = ""}

    }

    $AddStr = $AddStr -replace ("[^0-9a-zA-Z_]","")
    $AddStr = $AddStr -replace (" ","")
    

    if(!($FilestoRestoreList.Count)){

        $RestoreDBName = $FilestoRestoreList.database_name+$AddStr

    }else{

        $RestoreDBName = $FilestoRestoreList[0].database_name+$AddStr
    }

    $RestoreDBName = $RestoreDBName -replace("-","")

}
if($Error){

    Write-Warning "An error has occurred, try again."
    $SourceSQLCon.ConnectionContext.Disconnect();$DestSQLCon.ConnectionContext.Disconnect();break;

}else{

    #show prompt, if yes collect email data
    [boolean]$SendEmail = PromptForEmailProfile
}
if($Error){

    Write-Warning "An error has occurred, try again."
    $SourceSQLCon.ConnectionContext.Disconnect();$DestSQLCon.ConnectionContext.Disconnect();break;

}else{

   if($SendEmail -eq $true){

        Write-Host "`n`nThe default DB Mail server is:    $DefaultMailSQLServer"
        Write-Host "The default mail profile is:      $DefaultMailProfile"
        Write-Host "`nSelect the server and profile above if not sure which to select."

        $title = "Select SQL Server to send email"
        $message = "Database mail will be used to send results of the job, usaully use the source or production
SQL server to send the mail, but either can be used. Be sure DB mail is setup on the SQL instance selected."

        $Dflt = New-Object System.Management.Automation.Host.ChoiceDescription "&Defaults listed above", `
            "Job will run but no email will be sent."
        
        $Src = New-Object System.Management.Automation.Host.ChoiceDescription "&Source SQL Server: $SourceServer", `
            "Enter yes to enter profile name and recipients of completion email."

        $Dst = New-Object System.Management.Automation.Host.ChoiceDescription "&Use Dest SQL Server: $DestSQLsvr", `
            "Job will run but no email will be sent."

        $ManStr = New-Object System.Management.Automation.Host.ChoiceDescription "&Enter Server and Profile Manually", `
            "Enter DBMail server and Profile name."

        $options = [System.Management.Automation.Host.ChoiceDescription[]]($Dflt,$Src,$Dst,$ManStr)

        $result = $host.ui.PromptForChoice($title, $message, $options, 0) 

        switch ($result){

            0{$DBMailSrvr = $DefaultMailSQLServer; $DBMailProfile = $DefaultMailProfile}
            1{$DBMailSrvr = $SourceServer;
                $DBMailProfileOpt = Get-MailProfiles -ServerConnection $SourceSQLCon | Out-GridView -Title "Select DB Mail Profile to Send Email" -OutputMode Single;
                $DBMailProfile = $DBMailProfileOpt.profile_name}
            2{$DBMailSrvr = $DestSQLsvr
                $DBMailProfileOpt =  Get-MailProfiles -ServerConnection $DestSQLCon  | Out-GridView -Title "Select DB Mail Profile to Send Email" -OutputMode Single;
                $DBMailProfile = $DBMailProfileOpt.profile_name}
            3{$DBMailSrvr = Read-Host "Enter name of SQL DBMail server";$DBMailProfile = Read-Host "Enter SQL DBMail profile to send email";}

        }


        [string]$DBMailRecipients = Read-Host "Enter recipients of email, seperate multiple by semicolon"

    }
}
if($Error){

    Write-Warning "An error has occurred, try again."
    $SourceSQLCon.ConnectionContext.Disconnect();$DestSQLCon.ConnectionContext.Disconnect();break;

}else{

    #show prompt, if yes run restore
    PromptForConfirmation
}

#Disconnect from SQL Servers
$SourceSQLCon.ConnectionContext.Disconnect()
$DestSQLCon.ConnectionContext.Disconnect()
