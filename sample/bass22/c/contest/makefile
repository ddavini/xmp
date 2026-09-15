CC = gcc
RES = windres
FLAGS = -O2 -s -I..
BASSLIB = ..\bass.lib
OUTDIR = ..\bin

contest:
	$(CC) $(FLAGS) $@.c $(BASSLIB) -o $(OUTDIR)\$@.exe
