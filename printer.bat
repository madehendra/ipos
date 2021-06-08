@ECHO OFF
ECHO.
ECHO This is a batch file
ECHO.
net use lpt1: /delete
NET USE LPT1: \\127.0.0.1\pos PERSISTENT:NO
echo aaaaaaaaaaggggggggggaaaaaaaaaaggggggggggaaaaaaaaaagggggggggg > lpt1
echo aaaaaaaaaaggggggggggaaaaaaaaaaggggggggggaaaaaaaaaagggggggggg > lpt1
echo aaaaaaaaaaggggggggggaaaaaaaaaaggggggggggaaaaaaaaaagggggggggg > lpt1
echo aaaaaaaaaaggggggggggaaaaaaaaaaggggggggggaaaaaaaaaagggggggggg > lpt1
net use
PAUSE
