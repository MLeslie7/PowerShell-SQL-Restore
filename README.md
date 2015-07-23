# PowerShell-SQL-Restore
Restore SQL databases to different servers, works well with Ola Hallengren backup scripts especially when saving backup files to a network share and restoring onto different servers for dev-prod.
https://ola.hallengren.com/sql-server-backup.html

The script will run correctly as ‘administrator’ from a 2012,8.0 or newer machine with SQL Management Objects installed, you can query backups taken on a 2005 or newer SQL box and restore to a server running SQL 2005, 2008 or 2012, 2014 instance and save the files to a network share if restored on a SQL 2012 or newer instance. Use any machine with PowerShell 4.0 or higher and SQL Management Objects (MSQLSMS complete) installed.

Support for network path restores or local folder, Recovery-NoRecovery option, will NOT do a restore with Replace on purpose. Allows to rename the restored DB or use original name. Will send an email when restore is completed, good for large DB restores.

Script is used in production with 150+ GB DBs running on a Availaability Group, does not join the restored DB to the group - I am thinking of adding that in a future update.


