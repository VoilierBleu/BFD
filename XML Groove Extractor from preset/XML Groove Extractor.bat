:: https://github.com/VoilierBleu/BFD : 11:50 vendredi 10 décembre 2021
@echo off
	setlocal enabledelayedexpansion
	set $loop_counter=0
	:Loop
		set /a $loop_counter+=1
		title [!$loop_counter!] %~dpn0.vbs
		cls
		for /f "delims=" %%i in ('grep -i .bfd3$ "D:\AUDIO\Librairies\Drums\BFD3\locate"') do (
           title [%%i]
            echo Processing [%%i]
            cscript //nologo "%~dpn0.vbs" "%%i"
            echo Done
            echo.
		)
		pause
	goto :Loop