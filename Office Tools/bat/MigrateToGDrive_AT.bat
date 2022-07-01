@echo off
cd ..
set /p TsrcPath=<conf/cli_backup/cliSrcPath
set /p TdestPath=<conf/cli_backup/cliDestPath
set /p Tprocessor=<conf/cli_backup/cliProcessor
if exist "%TsrcPath%" (
	if exist "%TdestPath%" (
		call robocopy "%TsrcPath%" "%TdestPath%" * /S /DCOPY:DAT /COMPRESS /NOOFFLOAD /MT:%Tprocessor% /LOG+:log/robolog /TS /FP /TEE /V /ETA /R:5
		if %ERRORLEVEL% == 0 goto :next
		if %ERRORLEVEL% == 4 goto :err4
		if %ERRORLEVEL% == 8 goto :err8
		if %ERRORLEVEL% == 16 goto :err16
		echo err>> "log/lastResult"
		echo # Office Tools v1.2 >> "log/err"
		echo Backup Result			: Error >> "log/err"
		echo Reason			: Unspecified Error: %errorlevel% >> "log/err"
		echo Source Path			: %TsrcPath% >> "log/err"
		echo Destination Path		: %TdestPath% >> "log/err"
		echo Backup Time			: %date% - %time% >> "log/err"
		echo Backup Pref			: Anytime >> "log/err"
		echo Backup Type			: Auto >> "log/err"
		echo. >> "log/err"
		goto :endofscript

		:err4
		echo err>> "log/lastResult"
		echo # Office Tools v1.2 >> "log/err" 
		echo Backup Result			: Error >> "log/err"
		echo Reason			: Some mismatched files or directories were detected ! >> "log/err"
		echo Some mismatched files or directories were detected ! >> "log/lastErr"
		echo Source Path			: %TsrcPath% >> "log/err" 
		echo Destination Path		: %TdestPath% >> "log/err" 
		echo Backup Time			: %date% - %time% >> "log/err"
		echo Backup Pref			: Anytime >> "log/err"
		echo Backup Type			: Auto >> "log/err"
		echo. >> "log/err"
		goto :endofscript

		:err8
		echo err>> "log/lastResult"
		echo # Office Tools v1.2 >> "log/err" 
		echo Backup Result			: Error >> "log/err"
		echo Reason			: Some files or directories could not be copied !>> "log/err"
		echo Some files or directories could not be copied ! >> "log/lastErr"
		echo Source Path			: %TsrcPath%
		echo Destination Path		: %TdestPath%
		echo Backup Time			: %date% - %time% >> "log/err"
		echo Backup Pref			: Anytime >> "log/err"
		echo Backup Type			: Auto >> "log/err"
		echo. >> "log/err"
		goto :endofscript

		:err16
		echo err>> "log/lastResult"
		echo # Office Tools v1.2 >> "log/err" 
		echo Backup Result			: Error >> "log/err"
		echo Reason			: Serious error. Robocopy did not copy any files ! >> "log/err"
		echo Serious error. Robocopy did not copy any files ! >> "log/lastErr"
		echo Source Path			: %TsrcPath%
		echo Destination Path		: %TdestPath%
		echo Backup Time			: %date% - %time% >> "log/err"
		echo Backup Pref			: Anytime >> "log/err"
		echo Backup Type			: Auto >> "log/err"
		echo. >> "log/err"
		goto :endofscript

		:next
		echo success>> "log/lastResult"
		echo # Office Tools v1.2 >> "log/log"
		echo Backup Result			: Success >> "log/log"
		echo Source Path			: %TsrcPath% >> "log/log"
		echo Destination Path		: %TdestPath% >> "log/log"
		echo Backup Time			: %date% - %time% >> "log/log"
		echo Backup Pref			: Anytime >> "log/log"
		echo Backup Type			: Auto >> "log/log"
		echo. >> "log/log"
		goto :endofscript
	) else (
		echo err>> "log/lastResult"
		echo # Office Tools v1.2 >> "log/err" 
		echo Backup Result				: Error >> "log/err"
		echo Reason					: Destination path not exists >> "log/err"
		echo Source Path				: %TsrcPath% >> "log/err"
		echo Destination Path			: %TdestPath% >> "log/err"
		echo Backup Time				: %date% - %time% >> "log/err"
		echo Backup Pref				: Anytime >> "log/err"
		echo Backup Type				: Auto >> "log/err"
		echo. >> "log/err"
		goto :endofscript
	)
) else (
	echo err>> "log/lastResult"
	echo # Office Tools v1.2 >> "log/err" 
	echo Backup Result			: Error >> "log/err"
	echo Reason			: Source path not exists >> "log/err"
	echo Source Path			: %TsrcPath% >> "log/err"
	echo Destination Path		: %TdestPath% >> "log/err"
	echo Backup Time			: %date% - %time% >> "log/err"
	echo Backup Pref			: Anytime >> "log/err"
	echo Backup Type			: Auto >> "log/err"
	echo. >> "log/err"
	goto :endofscript
)

:endofscript
exit