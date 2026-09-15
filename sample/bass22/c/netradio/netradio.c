// BASS Internet radio example, copyright (c) 2002-2005 Ian Luck.

#include <windows.h>
#include <process.h>
#include <stdio.h>
#include "bass.h"


HWND win=NULL;
DWORD cthread=0;

HSTREAM chan;

char *urls[10]={ // stream URLs
	"http://64.236.34.196/stream/1048","http://205.188.234.129:8024",
	"http://64.236.34.97/stream/1006","http://206.98.167.99:8406",
	"http://160.79.1.141:8000","http://streams.riverhosting.net:8012",
	"http://205.188.234.4:8016","http://205.188.234.4:8014",
	"http://207.200.96.225:8010","http://64.202.98.91:8082"
};

// display error messages
void Error(char *es)
{
	char mes[200];
	sprintf(mes,"%s\n(error code: %d)",es,BASS_ErrorGetCode());
	MessageBox(win,mes,"Error",0);
}

#define MESS(id,m,w,l) SendDlgItemMessage(win,id,m,(WPARAM)w,(LPARAM)l)

// update stream title from metadata
void DoMeta(char *meta)
{
	char *p;
	if (meta && (p=strstr(meta,"StreamTitle='"))) {
		p=strdup(p+13);
		strchr(p,';')[-1]=0;
		MESS(30,WM_SETTEXT,0,p);
		free(p);
	}
}

void CALLBACK MetaSync(HSYNC handle, DWORD channel, DWORD data, DWORD user)
{
	DoMeta((char*)data);
}

void CALLBACK StatusProc(void *buffer,DWORD length,DWORD user)
{
	if (buffer && !length)
		MESS(32,WM_SETTEXT,0,buffer); // display connection status
}

void __cdecl OpenURL(char *url)
{
	BASS_StreamFree(chan); // close old stream
	MESS(31,WM_SETTEXT,0,"connecting...");
	MESS(30,WM_SETTEXT,0,"");
	MESS(32,WM_SETTEXT,0,"");
	if (!(chan=BASS_StreamCreateURL(url,0,BASS_STREAM_STATUS,StatusProc,0))) {
		MESS(31,WM_SETTEXT,0,"not playing");
		Error("Can't play the stream");
	} else {
		while (1) { // pre-buffer...
			char text[20];
			DWORD progress,len=BASS_StreamGetFilePosition(chan,BASS_FILEPOS_END);
			if (len==-1) goto done; // something's gone wrong! (eg. BASS_Free called)
			progress=(BASS_StreamGetFilePosition(chan,BASS_FILEPOS_DOWNLOAD)
				-BASS_StreamGetFilePosition(chan,BASS_FILEPOS_CURRENT))*100/len; // percentage of buffer filled
			if (progress>75) break; // over 75% full, enough
			sprintf(text,"buffering... %d%%",progress);
			MESS(31,WM_SETTEXT,0,text);
			Sleep(50);
		}
		{ // get the broadcast name and bitrate
			char *icy=BASS_StreamGetTags(chan,BASS_TAG_ICY);
			if (icy)
				for (;*icy;icy+=strlen(icy)+1) {
					if (!memcmp(icy,"icy-name:",9))
						MESS(31,WM_SETTEXT,0,icy+9);
					if (!memcmp(icy,"icy-br:",7)) {
						char br[30]="bitrate: ";
						strcat(br,icy+7);
						MESS(32,WM_SETTEXT,0,br);
					}
				}
			else
				MESS(31,WM_SETTEXT,0,"");
		}
		// get the stream title and set sync for subsequent titles
		DoMeta(BASS_StreamGetTags(chan,BASS_TAG_META));
		BASS_ChannelSetSync(chan,BASS_SYNC_META,0,&MetaSync,0);
		// play it!
		BASS_ChannelPlay(chan,FALSE);
	}
done:
	cthread=0;
}

BOOL CALLBACK dialogproc(HWND h,UINT m,WPARAM w,LPARAM l)
{
	switch (m) {
		case WM_COMMAND:
			switch (LOWORD(w)) {
				case IDCANCEL:
					DestroyWindow(h);
					return 1;
				default:
					if (LOWORD(w)>=10 && LOWORD(w)<20) {
						if (cthread) { // already connecting
							MessageBeep(0);
							break;
						}
						// open URL in a new thread (so that main thread is free)
						cthread=_beginthread(OpenURL,0,urls[LOWORD(w)-10]);
					}
			}
			break;

		case WM_INITDIALOG:
			win=h;
			// setup output device
			if (!BASS_Init(-1,44100,0,win,NULL)) {
				Error("Can't initialize device");
				DestroyWindow(win);
			}
			BASS_SetConfig(BASS_CONFIG_NET_PREBUF,0); // minimize automatic pre-buffering, so we can do it (and display it) instead
			return 1;

		case WM_DESTROY:
			BASS_Free();
			break;
	}
	return 0;
}

int PASCAL WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance,LPSTR lpCmdLine, int nCmdShow)
{
	// check that BASS 2.2 was loaded
	if (BASS_GetVersion()!=MAKELONG(2,2)) {
		MessageBox(0,"BASS version 2.2 was not loaded","Incorrect BASS.DLL",0);
		return 0;
	}

	// display the window
	DialogBox(hInstance,MAKEINTRESOURCE(1000),0,&dialogproc);

	return 0;
}
