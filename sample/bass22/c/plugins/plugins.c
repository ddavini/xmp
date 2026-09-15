// BASS plugin test, copyright (c) 2005 Ian Luck.

#include <windows.h>
#include <stdio.h>
#include <math.h>
#include "bass.h"
#include <commctrl.h>

HWND win=NULL;

DWORD chan;	// the channel

OPENFILENAME ofn;
char path[MAX_PATH];

// display error messages
void Error(char *es)
{
	char mes[200];
	sprintf(mes,"%s\n(error code: %d)",es,BASS_ErrorGetCode());
	MessageBox(win,mes,"Error",0);
}

// translate a CTYPE value to text
char *GetCTypeString(DWORD ctype)
{
	if (ctype==BASS_CTYPE_SAMPLE) return "sample";
	if (ctype==BASS_CTYPE_RECORD) return "recording";
	if (ctype&BASS_CTYPE_MUSIC_MOD) { // MOD music type...
		if (ctype&BASS_CTYPE_MUSIC_MO3) return "MO3";
		if (ctype==BASS_CTYPE_MUSIC_MTM) return "MTM";
		if (ctype==BASS_CTYPE_MUSIC_S3M) return "S3M";
		if (ctype==BASS_CTYPE_MUSIC_XM) return "XM";
		if (ctype==BASS_CTYPE_MUSIC_IT) return "IT";
		return "MOD";
	}
	if (ctype&BASS_CTYPE_STREAM) { // stream type...
		if (ctype==BASS_CTYPE_STREAM) return "custom stream";
		if (ctype==BASS_CTYPE_STREAM_WAV) return "WAV";
		if (ctype==BASS_CTYPE_STREAM_OGG) return "OGG";
		if (ctype==BASS_CTYPE_STREAM_MP1) return "MP1";
		if (ctype==BASS_CTYPE_STREAM_MP2) return "MP2";
		if (ctype==BASS_CTYPE_STREAM_MP3) return "MP3";
		if (ctype==BASS_CTYPE_STREAM_AIFF) return "AIFF";
		// check add-ons...
#define BASS_CTYPE_STREAM_CD	0x10200
		if (ctype==BASS_CTYPE_STREAM_CD) return "CDA";
#define BASS_CTYPE_STREAM_WMA	0x10300
		if (ctype==BASS_CTYPE_STREAM_WMA) return "WMA";
#define BASS_CTYPE_STREAM_FLAC	0x10900
		if (ctype==BASS_CTYPE_STREAM_FLAC) return "FLAC";
#define BASS_CTYPE_STREAM_WV	0x10500
#define BASS_CTYPE_STREAM_WV_LH	0x10503
		if (ctype>=BASS_CTYPE_STREAM_WV && ctype<=BASS_CTYPE_STREAM_WV_LH) return "Wavpack";
#define BASS_CTYPE_STREAM_OFR	0x10600
		if (ctype==BASS_CTYPE_STREAM_OFR) return "Optimfrog";
#define BASS_CTYPE_STREAM_APE	0x10700
		if (ctype==BASS_CTYPE_STREAM_APE) return "APE";
#define BASS_CTYPE_STREAM_MPC	0x10a00
		if (ctype==BASS_CTYPE_STREAM_MPC) return "MPC";
#define BASS_CTYPE_STREAM_AAC	0x10b00
		if (ctype==BASS_CTYPE_STREAM_AAC) return "AAC";
#define BASS_CTYPE_STREAM_MP4	0x10b01
		if (ctype==BASS_CTYPE_STREAM_MP4) return "MP4";
#define BASS_CTYPE_STREAM_SPX	0x10c00
		if (ctype==BASS_CTYPE_STREAM_SPX) return "Speex";
#define BASS_CTYPE_STREAM_ALAC	0x10e00
		if (ctype==BASS_CTYPE_STREAM_ALAC) return "ALAC";
#define BASS_CTYPE_STREAM_TTA	0x10f00
		if (ctype==BASS_CTYPE_STREAM_TTA) return "TTA";
#define BASS_CTYPE_STREAM_AC3	0x11000
		if (ctype==BASS_CTYPE_STREAM_AC3) return "AC3";
		return "unknown add-on";
	}
	return "?";
}

#define MESS(id,m,w,l) SendDlgItemMessage(win,id,m,(WPARAM)w,(LPARAM)l)

BOOL CALLBACK dialogproc(HWND h,UINT m,WPARAM w,LPARAM l)
{
	switch (m) {
		case WM_COMMAND:
			switch (LOWORD(w)) {
				case IDCANCEL:
					DestroyWindow(h);
					break;
				case 10:
					{
						BASS_CHANNELINFO info;
						char file[MAX_PATH]="";
						ofn.lpstrFilter="All files\0*.*\0\0";
						ofn.lpstrFile=file;
						if (GetOpenFileName(&ofn)) {
							memcpy(path,file,ofn.nFileOffset);
							path[ofn.nFileOffset-1]=0;
							// free the old stream
							BASS_StreamFree(chan);
							if (!(chan=BASS_StreamCreateFile(FALSE,file,0,0,BASS_SAMPLE_LOOP))) {
								// it ain't playable
								MESS(10,WM_SETTEXT,0,"click here to open a file...");
								MESS(11,WM_SETTEXT,0,"");
								Error("Can't play the file");
								break;
							}
							BASS_ChannelGetInfo(chan,&info);
							MESS(10,WM_SETTEXT,0,file);
							sprintf(file,"channel type = %x (%s)",info.ctype,GetCTypeString(info.ctype));
							MESS(11,WM_SETTEXT,0,file);
							{
								DWORD time=BASS_ChannelBytes2Seconds(chan,BASS_ChannelGetLength(chan));
								MESS(12,TBM_SETRANGE,1,MAKELONG(0,time));
							}
							BASS_ChannelPlay(chan,FALSE);
						}
					}
					break;
			}
			break;

		case WM_HSCROLL:
			if (l && LOWORD(w)!=SB_THUMBPOSITION && LOWORD(w)!=SB_ENDSCROLL) { // set the position
				int pos=SendMessage((HWND)l,TBM_GETPOS,0,0);
				BASS_ChannelSetPosition(chan,BASS_ChannelSeconds2Bytes(chan,pos));
			}
			break;

		case WM_TIMER:
			MESS(12,TBM_SETPOS,1,(DWORD)BASS_ChannelBytes2Seconds(chan,BASS_ChannelGetPosition(chan))); // update position
			break;

		case WM_INITDIALOG:
			win=h;
			GetCurrentDirectory(MAX_PATH,path);
			memset(&ofn,0,sizeof(ofn));
			ofn.lStructSize=sizeof(ofn);
			ofn.hwndOwner=h;
			ofn.nMaxFile=MAX_PATH;
			ofn.lpstrInitialDir=path;
			ofn.Flags=OFN_HIDEREADONLY|OFN_EXPLORER;
			// initialize default output device
			if (!BASS_Init(-1,44100,0,win,NULL)) {
				Error("Can't initialize device");
				DestroyWindow(win);
				break;
			}
			{ // look for plugins (in the executable's directory)
				WIN32_FIND_DATA fd;
				HANDLE fh;
				char path[MAX_PATH];
				GetModuleFileName(0,path,sizeof(path));
				strcpy(strrchr(path,'\\')+1,"bass*.dll");
				fh=FindFirstFile(path,&fd);
				if (fh!=INVALID_HANDLE_VALUE) {
					do {
						if (BASS_PluginLoad(fd.cFileName)) // plugin loaded...
							MESS(20,LB_ADDSTRING,0,fd.cFileName); //  add it to the list
					} while (FindNextFile(fh,&fd));
					FindClose(fh);
				}
				if (!MESS(20,LB_GETCOUNT,0,0)) // no plugins...
					MESS(20,LB_ADDSTRING,0,"no plugins - visit the BASS webpage to get some");
			}
			SetTimer(h,0,500,0); // timer to update the position
			return 1;

		case WM_DESTROY:
			// "free" the output device and all plugins
			BASS_Free();
			BASS_PluginFree(0);
			break;
	}
	return 0;
}

int PASCAL WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance,LPSTR lpCmdLine, int nCmdShow)
{
	// Check that BASS 2.2 was loaded
	if (BASS_GetVersion()!=MAKELONG(2,2)) {
		MessageBox(0,"BASS version 2.2 was not loaded","Incorrect BASS.DLL",0);
		return 0;
	}

	DialogBox(hInstance,(char*)1000,0,&dialogproc);

	return 0;
}
