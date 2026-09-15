CC = gcc
RES = windres
FLAGS = -O2 -s -I..
BASSLIB = ..\bass.lib
OUTDIR = ..\bin

livespec:
	$(CC) $(FLAGS) $@.c $(BASSLIB) -lgdi32 -lwinmm -mwindows -o $(OUTDIR)\$@.exe
