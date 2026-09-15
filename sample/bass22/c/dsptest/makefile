CC = gcc
RES = windres
FLAGS = -O2 -s -I..
BASSLIB = ..\bass.lib
OUTDIR = ..\bin

dsptest:
	$(RES) -i $@.rc -o rsrc.obj
	$(CC) $(FLAGS) $@.c rsrc.obj $(BASSLIB) -lcomdlg32 -mwindows -o $(OUTDIR)\$@.exe
