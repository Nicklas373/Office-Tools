@echo off
set /p TinstPath=<conf\adv_backup\advInstPath
set /p TsrcPath=<conf\adv_backup\advSrcPath
set /p TdestPath=<conf\adv_backup\advDestPath
set /p TcompLevel=<conf\adv_backup\advCompLvl
set /p TcompType=<conf\adv_backup\advCompType
set /p TrndmStrg=<conf\adv_backup\advRandomStrg
set /p TcompExt=<conf\adv_backup\advCompExt
set /p TencType=<conf\adv_backup\advEncType
for /f "skip=1" %%x in ('wmic os get localdatetime') do if not defined MyDate set MyDate=%%x
for /f %%x in ('wmic path win32_localtime get /format:list ^| findstr "="') do set %%x
set fmonth=00%Month%
set fday=00%Day%
set today=%fmonth:~-2%-%fday:~-2%-%Year%
if exist "%TsrcPath%" (
	if exist "%TdestPath%" (
		call "%TinstPath%\7z\7za" a %TcompType% %TcompLvl% OfficeTools_%today%_%TrndmStrg%.%TcompExt% "%TsrcPath%\*"
		if %ERRORLEVEL% == 0 goto :next
		echo err>> "log/lastResult"
		echo. >> "log/adverr"
		echo # Office Tools v1.2 >> "log/adverr"
		echo Compress Result				: Error >> "log/adverr"
		echo Compress Time				: %date% - %time% >> "log/adverr"
		echo Compress Filename			: OfficeTools_%today%_%TrndmStrg%.%TcompExt% >> "log/adverr"
		echo Compress Location			: %TdestPath% >> "log/adverr"
		echo Error Reason			: Unspecified Error Code: %errorlevel% >> "log/adverr"
		echo Encryption Method			: %TencType% >> "log/adverr"
		echo. >> "log/adverr"
		goto :endofscript
		
		:next
		move "*.%TcompExt%" "%TdestPath%"
		echo. >> "log/advlog"
		echo success>> "log/lastResult"
		echo # Office Tools v1.2 >> "log/advlog"
		echo Compress Result				: Success >> "log/advlog"
		echo Compress Time				: %date% - %time% >> "log/advlog"
		echo Compress Filename			: OfficeTools_%today%_%TrndmStrg%.%TcompExt% >> "log/advlog"
		echo Compress Location			: %TdestPath% >> "log/advlog"
		echo Encryption Method			: %TencType% >> "log/advlog"
		echo. >> "log/advlog"
		goto :endofscript
	) else (
		echo. >> "log/adverr"
		echo # Office Tools v1.2 >> "log/adverr"
		echo Compress Result				: Error >> "log/adverr"
		echo Compress Time				: %date% - %time% >> "log/adverr"
		echo Compress Filename			: OfficeTools_%today%_%TrndmStrg%.%TcompExt% >> "log/adverr"
		echo Compress Location			: %TdestPath% >> "log/adverr"
		echo Error Reason			: Destination path not exist ! >> "log/adverr"
		echo Encryption Method			: %TencType% >> "log/adverr"
		echo. >> "log/adverr"
		goto :endofscript
	)
) else (
	echo. >> "log/adverr"
	echo # Office Tools v1.2 >> "log/adverr"
	echo Compress Result				: Error >> "log/adverr"
	echo Compress Time				: %date% - %time% >> "log/adverr"
	echo Compress Filename			: OfficeTools_%today%_%TrndmStrg%.%TcompExt% >> "log/adverr"
	echo Compress Location			: %TdestPath% >> "log/adverr"
	echo Error Reason			: Source path not exists ! >> "log/adverr"
	echo Encryption Method			: %TencType% >> "log/adverr"
	echo. >> "log/adverr"
	goto :endofscript
)

:endofscript
exit