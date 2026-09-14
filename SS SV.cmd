@echo off
chcp 65001 > nul
setlocal EnableDelayedExpansion

:: =====================================================================
::  TV Control Master V3
::  ปรับปรุงจาก V2 : ยิงขนาน + ตัด IP ซ้ำ + เช็คผลจริง + เก็บ log
::
::  วิธีใช้
::    ON_Off_DS_All_V3.cmd                 -> เปิดเมนู
::    ON_Off_DS_All_V3.cmd off 1 3 19      -> สั่งปิดโซน 1,3,19 แล้วจบ
::    ON_Off_DS_All_V3.cmd off all         -> ปิดทุกโซน (ใส่ Task Scheduler ได้)
::    ON_Off_DS_All_V3.cmd on 13           -> เปิดโซน 13
::    ON_Off_DS_All_V3.cmd restart all     -> รีสตาร์ททุกโซน
:: =====================================================================

:: ---------- ตั้งค่า ----------
set "ENDPOINT=/signber-control"
:: ยิงพร้อมกันกี่เครื่องต่อรอบ (ลดลงถ้าเน็ตตัน)
set "PARALLEL=30"
:: วินาที รอเชื่อมต่อ
set "CONNECT_TIMEOUT=2"
:: วินาที รอทั้งคำสั่ง (curl จะตัดเองเมื่อครบ)
set "MAX_TIME=4"

set "action=tv-off.php"
set "mode_text=ปิด [OFF]"

:: ---------- เตรียมโฟลเดอร์ทำงาน ----------
set "WORKDIR=%~dp0"
set "LOGDIR=%WORKDIR%logs"
set "TMPDIR=%TEMP%\tvctl_%RANDOM%"
if not exist "%LOGDIR%" mkdir "%LOGDIR%" > nul 2>&1
mkdir "%TMPDIR%" > nul 2>&1
set "QUEUE=%TMPDIR%\queue.txt"
set "WORKER=%TMPDIR%\_worker.cmd"

:: สร้างตัวช่วยยิง curl แบบเบื้องหลัง
> "%WORKER%" echo @echo off
>>"%WORKER%" echo curl -s -o nul -w "%%%%{http_code}" --connect-timeout %%2 --max-time %%3 "http://%%1%ENDPOINT%/%%4" ^> "%%5\%%1.res" 2^>nul

where curl > nul 2>&1
if errorlevel 1 (
    echo [ERROR] เครื่องนี้ไม่มี curl - ต้องเป็น Windows 10 1803 ขึ้นไป
    pause & goto :cleanup_exit
)

for /f "usebackq delims=" %%i in (`powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd"`) do set "TODAY=%%i"
set "LOGFILE=%LOGDIR%\tvctl_%TODAY%.log"

:: ---------- โหมดสั่งจาก command line ----------
if "%~1"=="" goto menu
if /I "%~1"=="off"     ( set "action=tv-off.php"     & set "mode_text=ปิด [OFF]" )
if /I "%~1"=="on"      ( set "action=tv-on.php"      & set "mode_text=เปิด [ON]" )
if /I "%~1"=="restart" ( set "action=tv-restart.php" & set "mode_text=รีสตาร์ท [RESTART]" )
set "choice=%*"
set "choice=!choice:%~1=!"
if "!choice: =!"=="" set "choice=all"
for /f "tokens=*" %%x in ("!choice!") do set "choice=%%x"
set "BATCHMODE=1"
goto execute

:: =====================================================================
:menu
cls
echo ==========================================================
echo               TV Control Master  V3
echo ==========================================================
echo Mode: *** %mode_text% ***    ^| ยิงขนานรอบละ %PARALLEL% เครื่อง
echo ----------------------------------------------------------
echo [M] สลับโหมด (ปิด - เปิด - รีสตาร์ท)
echo [P] ตั้งจำนวนยิงขนาน
echo ----------------------------------------------------------
echo [1] Signage Zone Information 1  (167)
echo [2] Signage Zone Meeting Room 1  (105)
echo [3] SV_Off V2 Meeting Room 1  (29)
echo [4] SV_Off V2 Meeting Room 2  (31)
echo [5] SV_Off V2.2  (39)
echo [6] SV_Off V2  (43)
echo [7] CA 5 6 7 8  (26)
echo [8] ห้องรับรองชั้น 4  (3)
echo [9] โรงอาหาร สผ ชั้น 1  (10)
echo [10] 2 ห้องกระทู้ใหม่  (8)
echo [11] ห้องจัดเลี้ยง  (17)
echo [12] ห้องรับรอง สส  (7)
echo [13] โรงอาหาร สว ชั้น 2  (4)
echo [14] ห้องทำงานจนท.กรรมาธิการ สว ชั้น 3  (4)
echo [15] ห้องรับรองพรรค สว ชั้น 2 (มีคนใช้)  (2)
echo [16] ห้องรับรองพรรค สว ชั้น 2  (9)
echo [17] ห้องรับอาหารชั้น 4 ห้องรับรอง  (2)
echo [18] ห้องประชุมงบประมาณ 14  (14)
echo [19] Kiosk  (50)
echo ----------------------------------------------------------
echo [99] สั่ง %mode_text% *** ทุกโซนพร้อมกัน ***
echo [0] ออกจากโปรแกรม
echo ==========================================================
echo พิมพ์หมายเลขโซน เว้นวรรคได้ เช่น 1 3 19  ^| M=สลับโหมด  P=ตั้งขนาน
set "choice="
set /p choice="Select Command: "

if not defined choice goto menu
if "%choice%"=="0" goto cleanup_exit
if /I "%choice%"=="m" goto toggle_mode
if /I "%choice%"=="p" goto set_parallel

:: ---------- ตั้งเวลาหน่วง ----------
echo.
set "delay=0"
set /p delay="ตั้งเวลาหน่วงก่อนสั่งงาน (วินาที) [Enter = สั่งทันที]: "
echo %delay%| findstr /r "^[0-9][0-9]*$" > nul || set "delay=0"
if not "%delay%"=="0" (
    echo.
    echo ระบบจะทำการ %mode_text% ในอีก %delay% วินาที...
    timeout /t %delay%
)

:: =====================================================================
:execute
echo.
echo กำลังรวบรวมรายการเครื่อง...

:: ล้างคิวและตัวนับ
break > "%QUEUE%"
del /q "%TMPDIR%\*.res" > nul 2>&1
for /f "delims==" %%v in ('set SEEN_ 2^>nul') do set "%%v="
set "QCOUNT=0"
set "ZONELIST="

if "%choice%"=="99" ( set "choice=all" )
if /I "%choice%"=="all" set "choice=1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19"

for %%i in (%choice%) do call :queue_zone %%i

if %QCOUNT%==0 (
    echo [!] ไม่มีเครื่องในคิว - ตรวจหมายเลขโซนอีกที
    if defined BATCHMODE goto cleanup_exit
    timeout /t 3 > nul & goto menu
)

echo โซนที่เลือก: !ZONELIST!
echo จำนวนเครื่อง (ตัดซ้ำแล้ว): %QCOUNT%
echo.
call :log "=== %mode_text% ^| zones:!ZONELIST! ^| targets:%QCOUNT% ==="
call :dispatch

:: ---------- สรุปผล ----------
set "OK=0"
set "FAIL=0"
set "FAILLIST="
for /f "usebackq delims=" %%a in ("%QUEUE%") do (
    set "CODE="
    if exist "%TMPDIR%\%%a.res" set /p CODE=<"%TMPDIR%\%%a.res"
    if "!CODE!"=="200" (
        set /a OK+=1
    ) else (
        set /a FAIL+=1
        set "FAILLIST=!FAILLIST! %%a"
        if "!CODE!"=="" ( call :log "  FAIL %%a no-response" ) else ( call :log "  FAIL %%a http=!CODE!" )
    )
)

echo.
echo ==========================================================
echo   %mode_text%  ->  สำเร็จ !OK! / ล้มเหลว !FAIL!   (ทั้งหมด %QCOUNT%)
if !FAIL! GTR 0 (
    echo ----------------------------------------------------------
    echo   เครื่องที่ไม่ตอบ:
    for %%f in (!FAILLIST!) do echo     - %%f
)
echo ----------------------------------------------------------
echo   log: %LOGFILE%
echo ==========================================================
call :log "SUMMARY ok=!OK! fail=!FAIL! total=%QCOUNT%"

if defined BATCHMODE (
    if !FAIL! GTR 0 ( set "RC=1" ) else ( set "RC=0" )
    goto cleanup_exit
)
echo.
pause
goto menu

:: =====================================================================
:: ยิงคำสั่งแบบขนานเป็นรอบ ๆ
:: =====================================================================
:dispatch
set /a SENT=0
set /a BATCHN=0
for /f "usebackq delims=" %%a in ("%QUEUE%") do (
    start "" /b "%WORKER%" %%a %CONNECT_TIMEOUT% %MAX_TIME% %action% "%TMPDIR%"
    set /a SENT+=1
    set /a BATCHN+=1
    if !BATCHN! GEQ %PARALLEL% (
        set /a BATCHN=0
        call :wait_batch !SENT!
        echo   ส่งแล้ว !SENT!/%QCOUNT%
    )
)
call :wait_batch %QCOUNT%
echo   ส่งครบ %QCOUNT%/%QCOUNT%
exit /b 0

:: รอจนผลลัพธ์ครบตามจำนวน หรือหมดเวลา
:wait_batch
set "WANT=%~1"
set /a SPIN=0
:wb_loop
set /a DONE=0
for %%f in ("%TMPDIR%\*.res") do set /a DONE+=1
if !DONE! GEQ %WANT% exit /b 0
set /a SPIN+=1
set /a LIMIT=%MAX_TIME%+3
if !SPIN! GEQ !LIMIT! exit /b 0
timeout /t 1 /nobreak > nul
goto wb_loop

:: =====================================================================
:: ใส่เครื่องลงคิว (ตัด IP ซ้ำอัตโนมัติ)
:: =====================================================================
:add
if defined SEEN_%~1 exit /b 0
set "SEEN_%~1=1"
>>"%QUEUE%" echo %~1
set /a QCOUNT+=1
exit /b 0

:: =====================================================================
:: แปลงหมายเลขโซน -> รายการเครื่อง
:: =====================================================================
:queue_zone
if "%~1"=="1" ( set "ZONELIST=!ZONELIST! [1]" & call :run_info1 )
if "%~1"=="2" ( set "ZONELIST=!ZONELIST! [2]" & call :run_mr1_sig )
if "%~1"=="3" ( set "ZONELIST=!ZONELIST! [3]" & call :run_mr1_sv2 )
if "%~1"=="4" ( set "ZONELIST=!ZONELIST! [4]" & call :run_mr2_sv2 )
if "%~1"=="5" ( set "ZONELIST=!ZONELIST! [5]" & call :run_sv22 )
if "%~1"=="6" ( set "ZONELIST=!ZONELIST! [6]" & call :run_sv2 )
if "%~1"=="7" ( set "ZONELIST=!ZONELIST! [7]" & call :run_ca5678 )
if "%~1"=="8" ( set "ZONELIST=!ZONELIST! [8]" & call :run_reception4 )
if "%~1"=="9" ( set "ZONELIST=!ZONELIST! [9]" & call :run_canteen )
if "%~1"=="10" ( set "ZONELIST=!ZONELIST! [10]" & call :run_newtopic )
if "%~1"=="11" ( set "ZONELIST=!ZONELIST! [11]" & call :run_banquet )
if "%~1"=="12" ( set "ZONELIST=!ZONELIST! [12]" & call :run_reception_mp )
if "%~1"=="13" ( set "ZONELIST=!ZONELIST! [13]" & call :run_canteen_sw2 )
if "%~1"=="14" ( set "ZONELIST=!ZONELIST! [14]" & call :run_office_sw3 )
if "%~1"=="15" ( set "ZONELIST=!ZONELIST! [15]" & call :run_party_sw2_used )
if "%~1"=="16" ( set "ZONELIST=!ZONELIST! [16]" & call :run_party_sw2 )
if "%~1"=="17" ( set "ZONELIST=!ZONELIST! [17]" & call :run_dining4 )
if "%~1"=="18" ( set "ZONELIST=!ZONELIST! [18]" & call :run_budget14 )
if "%~1"=="19" ( set "ZONELIST=!ZONELIST! [19]" & call :run_kiosk )
exit /b 0

:: =====================================================================
:toggle_mode
if "%action%"=="tv-off.php" (
    set "action=tv-on.php"
    set "mode_text=เปิด [ON]"
) else if "%action%"=="tv-on.php" (
    set "action=tv-restart.php"
    set "mode_text=รีสตาร์ท [RESTART]"
) else (
    set "action=tv-off.php"
    set "mode_text=ปิด [OFF]"
)
goto menu

:set_parallel
echo.
set /p PARALLEL="ยิงพร้อมกันกี่เครื่องต่อรอบ (แนะนำ 20-50): "
echo %PARALLEL%| findstr /r "^[1-9][0-9]*$" > nul || set "PARALLEL=30"
goto menu

:log
>>"%LOGFILE%" echo [%TODAY% %TIME:~0,8%] %~1
exit /b 0

:cleanup_exit
rd /s /q "%TMPDIR%" > nul 2>&1
if not defined RC set "RC=0"
exit /b %RC%

:: =====================================================================
:: ชุด IP แต่ละโซน  -  เพิ่ม/ลบเครื่องแก้ตรงนี้ที่เดียว
:: (IP ซ้ำข้ามโซนไม่เป็นไร ระบบตัดซ้ำให้ตอนเข้าคิว)
:: =====================================================================

:run_info1
:: [1] Signage Zone Information 1  -  167 เครื่อง
call :add 10.14.19.15
call :add 10.14.19.16
call :add 10.12.19.14
call :add 10.14.19.13
call :add 10.21.19.17
call :add 10.14.49.16
call :add 10.24.19.11
call :add 10.51.19.62
call :add 10.54.19.12
call :add 10.54.19.13
call :add 10.11.29.15
call :add 10.14.29.11
call :add 10.21.29.17
call :add 10.24.29.12
call :add 10.41.29.17
call :add 10.54.29.17
call :add 10.11.39.14
call :add 10.11.39.13
call :add 10.12.39.14
call :add 10.13.39.15
call :add 10.54.39.16
call :add 10.11.49.11
call :add 10.12.49.11
call :add 10.11.49.12
call :add 10.54.49.17
call :add 10.11.59.11
call :add 10.54.59.14
call :add 10.14.69.11
call :add 10.54.69.16
call :add 10.11.79.11
call :add 10.13.79.11
call :add 10.14.89.11
call :add 10.13.89.11
call :add 10.64.179.20
call :add 10.64.179.21
call :add 10.64.179.22
call :add 10.64.179.23
call :add 10.64.179.24
call :add 10.41.19.13
call :add 10.41.19.18
call :add 10.41.19.19
call :add 10.54.19.11
call :add 10.11.29.11
call :add 10.11.29.13
call :add 10.11.29.14
call :add 10.14.29.14
call :add 10.14.29.15
call :add 10.14.29.16
call :add 10.14.29.17
call :add 10.24.29.14
call :add 10.24.29.16
call :add 10.24.29.11
call :add 10.41.29.11
call :add 10.41.29.13
call :add 10.41.29.16
call :add 10.44.29.11
call :add 10.51.29.12
call :add 10.54.29.15
call :add 10.44.39.13
call :add 10.54.39.11
call :add 10.11.49.13
call :add 10.51.49.15
call :add 10.51.49.16
call :add 10.54.49.14
call :add 10.11.59.12
call :add 10.22.59.12
call :add 10.51.59.14
call :add 10.51.59.16
call :add 10.11.79.12
call :add 10.11.79.13
call :add 10.14.79.14
call :add 10.12.79.11
call :add 10.12.79.12
call :add 10.12.79.13
call :add 10.13.79.12
call :add 10.11.89.11
call :add 10.11.89.12
call :add 10.12.89.11
call :add 10.12.89.12
call :add 10.13.89.12
call :add 10.14.89.13
call :add 10.24.89.11
call :add 10.23.89.11
call :add 10.0.99.11
call :add 10.0.99.12
call :add 10.0.99.13
call :add 10.0.99.14
call :add 10.11.99.13
call :add 10.14.99.11
call :add 10.14.99.12
call :add 10.60.169.11
call :add 10.60.169.12
call :add 10.60.169.13
call :add 10.64.169.12
call :add 10.64.169.13
call :add 10.64.169.11
call :add 10.14.159.11
call :add 10.60.179.20
call :add 10.60.179.21
call :add 10.60.179.22
call :add 10.60.179.23
call :add 10.60.179.24
call :add 10.60.179.25
call :add 10.51.19.11
call :add 10.13.29.16
call :add 10.13.29.17
call :add 10.14.29.12
call :add 10.14.29.13
call :add 10.21.29.12
call :add 10.21.29.13
call :add 10.21.29.14
call :add 10.21.29.15
call :add 10.21.29.16
call :add 10.0.89.11
call :add 10.21.19.11
call :add 10.21.19.12
call :add 10.21.19.13
call :add 10.21.19.14
call :add 10.21.19.15
call :add 10.21.19.16
call :add 10.41.19.20
call :add 10.22.19.11
call :add 10.22.19.12
call :add 10.22.19.13
call :add 10.22.19.14
call :add 10.22.19.15
call :add 10.22.19.16
call :add 10.22.19.17
call :add 10.41.19.11
call :add 10.41.19.12
call :add 10.41.19.14
call :add 10.41.19.15
call :add 10.41.19.16
call :add 10.41.19.17
call :add 10.51.19.12
call :add 10.0.109.11
call :add 10.0.109.12
call :add 10.0.109.13
call :add 10.0.109.14
call :add 10.11.29.12
call :add 10.24.29.13
call :add 10.24.29.17
call :add 10.24.29.18
call :add 10.24.29.19
call :add 10.24.29.20
call :add 10.24.29.21
call :add 10.24.29.22
call :add 10.24.29.23
call :add 10.24.29.24
call :add 10.41.29.12
call :add 10.41.29.14
call :add 10.41.29.15
call :add 10.44.29.12
call :add 10.51.29.11
call :add 10.54.29.14
call :add 10.44.39.12
call :add 10.44.39.11
call :add 10.51.49.13
call :add 10.54.49.18
call :add 10.54.49.11
call :add 10.54.49.13
call :add 10.51.59.13
call :add 10.13.79.13
call :add 10.14.79.11
call :add 10.14.79.12
call :add 10.44.169.11
call :add 10.54.39.12
exit /b 0

:run_mr1_sig
:: [2] Signage Zone Meeting Room 1  -  105 เครื่อง
call :add 10.51.29.13
call :add 10.51.29.14
call :add 10.51.29.15
call :add 10.54.29.12
call :add 10.54.29.13
call :add 10.11.39.19
call :add 10.11.39.18
call :add 10.11.39.17
call :add 10.11.39.16
call :add 10.11.39.15
call :add 10.12.39.15
call :add 10.12.39.13
call :add 10.13.39.14
call :add 10.14.39.11
call :add 10.14.39.12
call :add 10.14.39.13
call :add 10.14.39.14
call :add 10.21.39.12
call :add 10.21.39.13
call :add 10.21.39.16
call :add 10.21.39.14
call :add 10.21.39.15
call :add 10.21.39.21
call :add 10.21.39.20
call :add 10.21.39.19
call :add 10.21.39.18
call :add 10.24.39.15
call :add 10.24.39.11
call :add 10.24.39.12
call :add 10.24.39.14
call :add 10.24.39.13
call :add 10.24.39.16
call :add 10.24.39.17
call :add 10.24.39.18
call :add 10.24.39.19
call :add 10.24.39.20
call :add 10.24.39.21
call :add 10.51.39.11
call :add 10.51.39.12
call :add 10.54.39.13
call :add 10.54.39.14
call :add 10.54.39.15
call :add 10.11.49.14
call :add 10.11.49.16
call :add 10.11.49.15
call :add 10.12.49.14
call :add 10.12.49.15
call :add 10.21.49.16
call :add 10.21.49.11
call :add 10.21.49.15
call :add 10.21.49.13
call :add 10.21.49.12
call :add 10.21.49.14
call :add 10.21.49.20
call :add 10.21.49.19
call :add 10.21.49.18
call :add 10.21.49.17
call :add 10.23.49.15
call :add 10.23.49.17
call :add 10.23.49.18
call :add 10.23.49.19
call :add 10.23.49.20
call :add 10.23.49.21
call :add 10.24.49.17
call :add 10.24.49.18
call :add 10.24.49.19
call :add 10.24.49.20
call :add 10.24.49.12
call :add 10.24.49.14
call :add 10.24.49.11
call :add 10.24.49.15
call :add 10.24.49.13
call :add 10.24.49.16
call :add 10.41.49.11
call :add 10.41.49.12
call :add 10.41.49.13
call :add 10.41.49.14
call :add 10.41.49.15
call :add 10.44.49.11
call :add 10.44.49.12
call :add 10.44.49.13
call :add 10.44.49.14
call :add 10.44.49.15
call :add 10.51.49.11
call :add 10.51.49.12
call :add 10.51.49.14
call :add 10.51.49.18
call :add 10.54.49.12
call :add 10.54.49.15
call :add 10.54.49.16
call :add 10.51.59.11
call :add 10.51.59.12
call :add 10.51.59.15
call :add 10.54.59.11
call :add 10.54.59.12
call :add 10.54.59.13
call :add 10.51.69.11
call :add 10.51.29.16
call :add 10.54.69.13
call :add 10.54.69.14
call :add 10.54.69.15
call :add 10.11.99.11
call :add 10.11.99.12
call :add 10.21.29.11
call :add 10.54.29.11
exit /b 0

:run_mr1_sv2
:: [3] SV_Off V2 Meeting Room 1  -  29 เครื่อง
call :add 10.43.29.11
call :add 10.53.29.14
call :add 10.12.39.18
call :add 10.12.39.17
call :add 10.12.39.16
call :add 10.13.39.11
call :add 10.13.39.12
call :add 10.13.39.13
call :add 10.22.39.13
call :add 10.22.39.14
call :add 10.22.39.15
call :add 10.22.39.16
call :add 10.22.39.17
call :add 10.22.39.21
call :add 10.22.39.20
call :add 10.22.39.19
call :add 10.22.39.18
call :add 10.22.39.11
call :add 10.22.39.12
call :add 10.23.39.13
call :add 10.23.39.11
call :add 10.23.39.12
call :add 10.23.39.20
call :add 10.23.39.21
call :add 10.23.39.22
call :add 10.23.39.23
call :add 10.23.39.19
call :add 10.23.39.18
call :add 10.23.39.17
exit /b 0

:run_mr2_sv2
:: [4] SV_Off V2 Meeting Room 2  -  31 เครื่อง
call :add 10.23.39.16
call :add 10.23.39.15
call :add 10.23.39.14
call :add 10.42.39.11
call :add 10.42.39.12
call :add 10.12.49.16
call :add 10.22.49.13
call :add 10.22.49.14
call :add 10.22.49.15
call :add 10.22.49.16
call :add 10.22.49.17
call :add 10.22.49.21
call :add 10.22.49.20
call :add 10.22.49.19
call :add 10.22.49.18
call :add 10.22.49.11
call :add 10.22.49.12
call :add 10.23.49.11
call :add 10.23.49.12
call :add 10.23.49.13
call :add 10.23.49.14
call :add 10.23.49.16
call :add 10.42.49.11
call :add 10.43.49.11
call :add 10.52.49.12
call :add 10.53.49.15
call :add 10.53.49.16
call :add 10.42.19.13
call :add 10.52.19.13
call :add 10.42.29.11
call :add 10.42.29.13
exit /b 0

:run_sv22
:: [5] SV_Off V2.2  -  39 เครื่อง
call :add 10.13.19.16
call :add 10.12.19.11
call :add 10.12.19.13
call :add 10.13.19.12
call :add 10.52.19.16
call :add 10.53.19.11
call :add 10.13.29.11
call :add 10.11.29.16
call :add 10.23.29.12
call :add 10.22.29.12
call :add 10.53.29.16
call :add 10.12.39.19
call :add 10.13.39.16
call :add 10.53.39.16
call :add 10.13.49.11
call :add 10.53.49.17
call :add 10.13.59.11
call :add 10.13.69.11
call :add 10.52.19.11
call :add 10.52.19.12
call :add 10.52.19.15
call :add 10.12.29.11
call :add 10.12.29.12
call :add 10.12.29.14
call :add 10.13.29.12
call :add 10.13.29.13
call :add 10.13.29.14
call :add 10.13.29.15
call :add 10.22.29.11
call :add 10.23.29.13
call :add 10.23.29.15
call :add 10.23.29.20
call :add 10.42.29.12
call :add 10.42.29.14
call :add 10.43.29.12
call :add 10.52.29.11
call :add 10.53.29.15
call :add 10.11.39.12
call :add 10.11.39.11
exit /b 0

:run_sv2
:: [6] SV_Off V2  -  43 เครื่อง
call :add 10.12.39.12
call :add 10.12.39.11
call :add 10.43.39.11
call :add 10.52.39.12
call :add 10.52.39.13
call :add 10.53.39.12
call :add 10.53.39.15
call :add 10.12.49.12
call :add 10.12.49.13
call :add 10.53.49.12
call :add 10.22.59.11
call :add 10.13.99.11
call :add 10.13.99.12
call :add 10.13.99.13
call :add 10.22.29.13
call :add 10.22.29.14
call :add 10.22.29.15
call :add 10.53.49.11
call :add 10.53.49.13
call :add 10.42.19.11
call :add 10.42.19.12
call :add 10.52.19.14
call :add 10.12.29.13
call :add 10.23.29.11
call :add 10.23.29.14
call :add 10.23.29.16
call :add 10.23.29.17
call :add 10.23.29.18
call :add 10.23.29.19
call :add 10.23.29.21
call :add 10.23.29.22
call :add 10.23.29.23
call :add 10.43.29.13
call :add 10.52.29.12
call :add 10.53.29.13
call :add 10.52.39.11
call :add 10.53.39.11
call :add 10.53.39.13
call :add 10.53.39.14
call :add 10.52.49.11
call :add 10.53.49.14
call :add 10.13.59.12
call :add 10.43.169.11
exit /b 0

:run_ca5678
:: [7] CA 5 6 7 8  -  26 เครื่อง
call :add 10.24.89.11
call :add 10.13.89.12
call :add 10.14.89.13
call :add 10.11.89.11
call :add 10.12.89.12
call :add 10.12.89.11
call :add 10.11.89.12
call :add 10.23.89.11
call :add 10.0.89.11
call :add 10.11.79.13
call :add 10.12.79.11
call :add 10.12.79.12
call :add 10.12.79.13
call :add 10.13.79.11
call :add 10.13.79.13
call :add 10.14.79.14
call :add 10.14.79.12
call :add 10.14.79.11
call :add 10.13.79.12
call :add 10.11.79.12
call :add 10.13.59.13
call :add 10.13.59.14
call :add 10.14.59.12
call :add 10.14.59.11
call :add 10.11.59.12
call :add 10.13.59.12
exit /b 0

:run_reception4
:: [8] ห้องรับรองชั้น 4  -  3 เครื่อง
call :add 10.14.49.14
call :add 10.13.49.17
call :add 10.11.49.13
exit /b 0

:run_canteen
:: [9] โรงอาหาร สผ ชั้น 1  -  10 เครื่อง
call :add 10.41.19.13
call :add 10.41.19.11
call :add 10.41.19.17
call :add 10.41.19.15
call :add 10.41.19.18
call :add 10.41.19.16
call :add 10.41.19.14
call :add 10.41.19.12
call :add 10.41.19.20
call :add 10.41.19.19
exit /b 0

:run_newtopic
:: [10] 2 ห้องกระทู้ใหม่  -  8 เครื่อง
call :add 10.14.29.14
call :add 10.14.29.17
call :add 10.14.29.16
call :add 10.14.29.15
call :add 10.13.29.13
call :add 10.13.29.12
call :add 10.13.29.15
call :add 10.13.29.14
exit /b 0

:run_banquet
:: [11] ห้องจัดเลี้ยง  -  17 เครื่อง
call :add 10.41.19.12
call :add 10.22.19.11
call :add 10.22.19.12
call :add 10.22.19.13
call :add 10.22.19.15
call :add 10.22.19.17
call :add 10.21.19.12
call :add 10.21.19.13
call :add 10.21.19.14
call :add 10.21.19.15
call :add 10.22.19.16
call :add 10.42.19.11
call :add 10.42.19.12
call :add 10.21.19.11
call :add 10.41.19.20
call :add 10.22.19.14
call :add 10.21.19.16
exit /b 0

:run_reception_mp
:: [12] ห้องรับรอง สส  -  7 เครื่อง
call :add 10.11.29.11
call :add 10.11.29.12
call :add 10.11.29.13
call :add 10.11.29.14
call :add 10.14.29.14
call :add 10.14.29.17
call :add 10.14.29.16
exit /b 0

:run_canteen_sw2
:: [13] โรงอาหาร สว ชั้น 2  -  4 เครื่อง
call :add 10.42.29.14
call :add 10.42.29.11
call :add 10.42.29.12
call :add 10.42.29.13
exit /b 0

:run_office_sw3
:: [14] ห้องทำงานจนท.กรรมาธิการ สว ชั้น 3  -  4 เครื่อง
call :add 10.12.39.11
call :add 10.12.39.12
call :add 10.11.39.12
call :add 10.11.39.11
exit /b 0

:run_party_sw2_used
:: [15] ห้องรับรองพรรค สว ชั้น 2 (มีคนใช้)  -  2 เครื่อง
call :add 10.23.29.11
call :add 10.23.29.22
exit /b 0

:run_party_sw2
:: [16] ห้องรับรองพรรค สว ชั้น 2  -  9 เครื่อง
call :add 10.23.29.21
call :add 10.23.29.17
call :add 10.23.29.13
call :add 10.23.29.18
call :add 10.23.29.23
call :add 10.23.29.19
call :add 10.23.29.14
call :add 10.23.29.16
call :add 10.23.29.15
exit /b 0

:run_dining4
:: [17] ห้องรับอาหารชั้น 4 ห้องรับรอง  -  2 เครื่อง
call :add 10.12.49.13
call :add 10.12.49.12
exit /b 0

:run_budget14
:: [18] ห้องประชุมงบประมาณ 14  -  14 เครื่อง
call :add 10.14.49.15
call :add 10.13.49.16
call :add 10.13.49.14
call :add 10.14.49.11
call :add 10.13.49.17
call :add 10.14.49.14
call :add 10.13.49.12
call :add 10.13.49.13
call :add 10.14.49.12
call :add 10.14.49.13
call :add 10.13.59.14
call :add 10.13.59.13
call :add 10.14.59.12
call :add 10.14.59.11
exit /b 0

:run_kiosk
:: [19] Kiosk  -  50 เครื่อง
call :add 10.14.19.15
call :add 10.14.19.16
call :add 10.14.19.13
call :add 10.21.19.17
call :add 10.24.19.11
call :add 10.13.19.16
call :add 10.12.19.11
call :add 10.12.19.14
call :add 10.12.19.13
call :add 10.13.19.12
call :add 10.23.19.11
call :add 10.51.19.62
call :add 10.54.19.13
call :add 10.52.19.16
call :add 10.53.19.11
call :add 10.11.29.15
call :add 10.14.29.11
call :add 10.21.29.17
call :add 10.24.29.12
call :add 10.13.29.11
call :add 10.11.29.16
call :add 10.23.29.12
call :add 10.22.29.12
call :add 10.41.29.17
call :add 10.54.29.17
call :add 10.53.29.16
call :add 10.11.39.13
call :add 10.11.39.14
call :add 10.13.39.15
call :add 10.12.39.14
call :add 10.13.39.16
call :add 10.12.39.19
call :add 10.54.39.16
call :add 10.53.39.16
call :add 10.11.49.11
call :add 10.11.49.12
call :add 10.12.49.11
call :add 10.13.49.11
call :add 10.54.49.17
call :add 10.53.49.17
call :add 10.11.59.11
call :add 10.13.59.11
call :add 10.54.59.14
call :add 10.14.69.11
call :add 10.13.69.11
call :add 10.54.69.16
call :add 10.11.79.11
call :add 10.13.79.11
call :add 10.14.89.11
call :add 10.13.89.11
exit /b 0