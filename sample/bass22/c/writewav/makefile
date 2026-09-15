CC = gcc
RES = windres
FLAGS = -O2 -s -I..
BASSLIB = ..\bass.lib
OUTDIR = ..\bin

writewav:
	$(CC) $(FLAGS) $@.c $(BASSLIB) -o $(OUTDIR)\$@.exe
