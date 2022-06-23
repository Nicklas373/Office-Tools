@echo off
set /p TsrcPath=<conf\res_backup\resSrcPath
set /p TdestPath=<conf\res_backup\resDestPath
set /p TtempKey=<conf\res_backup\resTempKey
set /p TinstPath=<conf\res_backup\resInstPath
set /p TdecMtd=<conf\res_backup\resDecType
set /p TlogRes=<log\lastResult
if exist "%TsrcPath%" (
	if exist "%TdestPath%" (
		call "%TinstPath%\7z\7za" x "%TsrcPath%" -p%TtempKey% -aoa -o"%TdestPath%" >> "conf\res_backup\resZipLog"
		if %ERRORLEVEL% == 0 goto :endofscript
		goto :endofscript
	) else (
		echo. >> "log/reserr"
		echo # Office Tools v1.2 >> "log/reserr"
		echo Extract Result				: Error >> "log/reserr"
		echo Extract Time				: %date% - %time% >> "log/reserr"
		echo Extract Filename				: %TsrcPath% >> "log/reserr"
		echo Extract Location				: %TdestPath% >> "log/reserr"
		echo Error Reason					: Destination path not exists ! >> "log/reserr"
		echo Decryption Method			: %TdecMtd% >> "log/reserr"
		goto :endofscript
	)
) else (
	echo. >> "log/reserr"
	echo # Office Tools v1.2 >> "log/reserr"
		echo Extract Result				: Error >> "log/reserr"
		echo Extract Time				: %date% - %time% >> "log/reserr"
		echo Extract Filename				: %TsrcPath% >> "log/reserr"
		echo Extract Location				: %TdestPath% >> "log/reserr"
		echo Error Reason					: Source file not exists ! >> "log/reserr"
		echo Decryption Method			: %TdecMtd% >> "log/reserr"
	goto :endofscript
)

:endofscript
exit