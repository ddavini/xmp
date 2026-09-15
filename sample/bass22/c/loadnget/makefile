CC = gcc
RES = windres
FLAGS = -O2 -s -I..
BASSLIB = ..\bass.lib
OUTDIR = ..\bin

loadnget:
	$(RES) -i $@.rc -o rsrc.obj
	$(CC) $(FLAGS) $@.c rsrc.obj $(BASSLIB) -lwinmm -o $(OUTDIR)\$@.exe
