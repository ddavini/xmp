CC = gcc
RES = windres
FLAGS = -O2 -s -I..
BASSLIB = ..\bass.lib
OUTDIR = ..\bin

livefx:
	$(RES) -i $@.rc -o rsrc.obj
	$(CC) $(FLAGS) $@.c rsrc.obj $(BASSLIB) -lcomctl32 -mwindows -o $(OUTDIR)\$@.exe
