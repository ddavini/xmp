CC = gcc
RES = windres
FLAGS = -O2 -s -I..
BASSLIB = ..\bass.lib
OUTDIR = ..\bin

spectrum:
	$(CC) $(FLAGS) $@.c $(BASSLIB) -lcomdlg32 -lgdi32 -lwinmm -mwindows -o $(OUTDIR)\$@.exe
