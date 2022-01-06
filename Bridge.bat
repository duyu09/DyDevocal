::Copyright (c) 2021 duyu.
set a=%1
set b=%a:~1,-1%
lame.exe --decode "%b%" "%b%__WAVfile.wav"
Dy_DeCore.exe "%b%__WAVfile.wav"
del /Q /F "%b%__WAVfile.wav"
echo %2_finished>%tmp%\%2_finished