CC = gcc
RES = windres
FLAGS = -O2 -s -I..
BASSLIB = ..\bass.lib
OUTDIR = ..\bin

custloop:
	$(CC) $(FLAGS) $@.c $(BASSLIB) -lcomdlg32 -lgdi32 -mwindows -o $(OUTDIR)\$@.exe
