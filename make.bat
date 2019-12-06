@echo off
set gccbase=G:\p_files\rtdk\mingw32-gcc5\bin
set gccbase=G:\p_files\rtdk\i686-8.1.0-win32-dwarf-rt_v6-rev0\mingw32\bin
set fbcbase=G:\p_files\rtdk\FBC

set opts=-std=c99 -ggdb3 -Wall -Wextra
set opts=-std=c99 -Os -Wall -Wextra

set PATH=%PATH%;%gccbase%;%fbcbase%
del bin\comhelper.o
del bin\combase.o
del bin\libcomhelper.a

gcc -c src\comhelper.c -o bin\comhelper.o %opts% 2> err_chelper.log
gcc -c src\combase.c -o bin\combase.o %opts% 2> err_chelper.log
ar r bin\libcomhelper.a bin\comhelper.o bin\combase.o