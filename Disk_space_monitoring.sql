----+----#----+----#----+----#----+----#----+----#----+----#----+----#----+----#
--
-- Transact-SQL script for
--
-- Disk space monitoring for SQL Server 2019
--
-- A new table, 'master.dbo.Server_Diskinfo', is created, which includes information of :
--   Server name
--   Instance name
--   Datetime
--   Drive letter
--   Free space of drive
--   Total space of drive
--   Free space of drive [%]
--   Space of drive used by SQL
--
-- Created on 2022.6.17
--
-- Reference : https://github.com/michalsadowski/SQLBuild
--
----+----#----+----#----+----#----+----#----+----#----+----#----+----#----+----#

USE [master]
GO

-- Activate Ole (Object Linking and Embedding) Automation Procedures option to use COM object
sp_configure 'show advanced options', 1
GO
RECONFIGURE
GO
sp_configure 'Ole Automation Procedures', 1
GO
RECONFIGURE
GO

-- U stands for (User-defined) table as opposed to system table
-- N stands for uNicode (2 bytes) character/string literal.  Same results without the N
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Server_DiskInfo]') AND type in (N'U'))
DROP TABLE [dbo].[Server_DiskInfo]
GO

SET ANSI_NULLS ON
SET QUOTED_IDENTIFIER ON
SET ANSI_PADDING ON
GO

-- Create table on PRIMARY file group
CREATE TABLE [dbo].[Server_DiskInfo](
	[Server_Name] [varchar](100) NULL,
	[Instance_Name] [varchar](100) NULL,
	[Serv_Date] [varchar](25) NULL,
	[Drive_Ltr] [char](1) NULL,
	[Free_MB] [int] NULL,
	[Total_MB] [int] NULL,
	[Free_Percent] [decimal](18, 2) NULL,
	[SpacebySQL_MB] [int] NULL
) ON [PRIMARY]
GO

SET ANSI_PADDING OFF
GO

-- P stands for SQL Stored Procedure
-- PC stands for Assembly (CLR (Common Language Runtime)) stored-procedure
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_GetSrvDiskInfo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[sp_GetSrvDiskInfo]
GO

CREATE PROCEDURE [dbo].[sp_GetSrvDiskInfo]
AS
SET NOCOUNT ON
DELETE FROM master.dbo.Server_DiskInfo where Serv_Date <= GETDATE()

-- Local temporary table
-- sp_MSforeachdb procedure is an undocumented procedure that allows to run the same command against all databases
-- sysfiles contains one row for each file in a database. This system table is a virtual table; it cannot be updated or modified directly
CREATE TABLE #tmpSqlSpace (DBName Varchar(25),Location Varchar(60),Size Varchar(8),Device Varchar(30))
Exec sp_MSforeachdb 'Use [?] Insert into #tmpSqlSpace Select Convert(Varchar(25),DB_Name())''Database'',
Convert(Varchar(60),filename),Convert(Varchar(8),size/128)''Size in MB'',Convert(Varchar(30),Name) from sysfiles'

/*
-- ## test #tmpSqlSpace ##
SELECT * INTO [tempdb].[dbo].[check_tmpSqlSpace] FROM #tmpSqlSpace;
GO
EXEC [master].[dbo].[sp_GetSrvDiskInfo];
GO
-- ## ##
*/

DECLARE @ServName VARCHAR(100), @InstName VARCHAR(100), @ServDate datetime

-- MachineName : Windows computer name on which the server instance is running
-- ServerName : Both the Windows server and instance information associated with a specified instance of SQL Server
SELECT @ServName = RTRIM(CONVERT(char(30), SERVERPROPERTY('MachineName')))
SELECT @InstName = RTRIM(CONVERT(char(40), SERVERPROPERTY('ServerName'))) 

SELECT @ServDate = GetDate()

DECLARE @hr int, @fso int, @drive char(1), @odrive int, @TotalSize varchar(20) 
DECLARE @MB bigint ; SET @MB = 1048576

CREATE TABLE #drives (
drive char(1) PRIMARY KEY,
FreeSpace int NULL,
TotalSize int NULL,
SQLDriveSize int NULL)

-- Use xp_fixeddrives to Monitor Free Space
INSERT #drives(drive,FreeSpace) EXEC master.dbo.xp_fixeddrives

/*
-- ## test #drives ##
SELECT * INTO [tempdb].[dbo].[check_drives] FROM #drives;
GO
EXEC [master].[dbo].[sp_GetSrvDiskInfo];
GO
-- ## ##
*/

-- FileSystemObject object is an object that allows to work with drives, folders, files, and so on
EXEC @hr=sp_OACreate 'Scripting.FileSystemObject',@fso OUT,1

IF @hr <> 0 EXEC sp_OAGetErrorInfo @fso

DECLARE @SQLDrvSize int 

-- ##################################################
-- Get SQL drive size and total size for each drive
-- ##################################################

DECLARE dcur CURSOR LOCAL FAST_FORWARD
FOR SELECT drive from #drives
ORDER by drive

OPEN dcur
FETCH NEXT FROM dcur INTO @drive
WHILE @@FETCH_STATUS=0
BEGIN

Select @SQLDrvSize=sum(Convert(Int,Size)) from #tmpSqlSpace where Substring(Location,1,1)=@drive
Select @TotalSize=0

-- sp_OAMethod gets a unique ID for each volume attached to filesystem object
EXEC @hr = sp_OAMethod @fso,'GetDrive', @odrive OUT, @drive
IF @hr <> 0 EXEC sp_OAGetErrorInfo @fso

-- sp_OAGetProperty retrieves properties of each drive and filesystem
EXEC @hr = sp_OAGetProperty @odrive,'TotalSize', @TotalSize OUT
IF @hr <> 0 EXEC sp_OAGetErrorInfo @odrive

UPDATE #drives SET SQLDriveSize=@SQLDrvSize, TotalSize=@TotalSize/@MB WHERE drive=@drive  

FETCH NEXT FROM dcur INTO @drive
END

CLOSE dcur
DEALLOCATE dcur

-- ##################################################
-- ##################################################

EXEC @hr=sp_OADestroy @fso
IF @hr <> 0 EXEC sp_OAGetErrorInfo @fso

insert into master.dbo.Server_DiskInfo (Server_Name, Instance_Name, Serv_Date, Drive_Ltr, Free_MB, Total_MB, Free_Percent, SpacebySQL_MB) 

SELECT rtrim(CONVERT(char(30), SERVERPROPERTY('MachineName'))), RTRIM(CONVERT(char(40), SERVERPROPERTY('ServerName'))) , CAST(@ServDate AS nvarchar(30)), drive,
FreeSpace as 'Free(MB)', TotalSize as 'Total(MB)', CAST((FreeSpace/(TotalSize*1.0))*100.0 as int) as 'Free(%)', SQLDriveSize FROM #drives ORDER BY drive

DROP TABLE #drives
DROP table #tmpSqlSpace
GO

-- ## Test execution ##
EXEC [master].[dbo].[sp_GetSrvDiskInfo]
GO
-- ## ##
