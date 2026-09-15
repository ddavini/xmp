// BASS 3D Test, copyright (c) 1999-2005 Ian Luck.

#include <windows.h>
#include <commctrl.h>
#include <math.h>
#include "bass.h"
#include "3dtest.h"

HWND win=NULL;

/* channel (sample/music) info structure */
typedef struct {
	DWORD channel;			// the channel
	BASS_3DVECTOR pos,vel;	// position,velocity
	int dir;				// direction of the channel
} Channel;

Channel *chans=NULL;		// the channels
int chanc=0,chan=-1;		// number of channels, current channel
HBRUSH brush1,brush2;	// brushes

#define TIMERPERIOD	50		// timer period (ms)
#define MAXDIST		50		// maximum distance of the channels (m)
#define SPEED		12		// speed of the channels' movement (m/s)

/* Display error dialogs */
void Error(char *es)
{
	char mes[200];
	sprintf(mes,"%s\n(error code: %d)",es,BASS_ErrorGetCode());
	MessageBox(win,mes,"Error",0);
}

/* Messaging macros */
#define ITEM(id) GetDlgItem(win,id)
#define MESS(id,m,w,l) SendDlgItemMessage(win,id,m,(WPARAM)w,(LPARAM)l)
#define LM(m,w,l) MESS(ID_LIST,m,w,l)
#define EM(m,w,l) MESS(ID_EAX,m,w,l)

void CALLBACK Update(HWND win, UINT m, UINT i, DWORD t)
{
	HDC dc;
    RECT r;
	int c,x,y,cx,cy;
	int save;
	HPEN pen;

	win=ITEM(ID_DISPLAY);
	dc=GetDC(win);
	save=SaveDC(dc);
	GetClientRect(win,&r);
	cx=r.right/2;
	cy=r.bottom/2;

	/* Draw center circle */
	SelectObject(dc,GetStockObject(GRAY_BRUSH));
	Ellipse(dc,cx-4,cy-4,cx+4,cy+4);

	pen=CreatePen(PS_SOLID,2,GetSysColor(COLOR_BTNFACE));
	SelectObject(dc,pen);

	for (c=0;c<chanc;c++) {
		/* If the channel's playing then update it's position */
		if (BASS_ChannelIsActive(chans[c].channel)==BASS_ACTIVE_PLAYING) {
			/* Check if channel has reached the max distance */
			if (chans[c].pos.z>=MAXDIST || chans[c].pos.z<=-MAXDIST)
				chans[c].vel.z=-chans[c].vel.z;
			if (chans[c].pos.x>=MAXDIST || chans[c].pos.x<=-MAXDIST)
				chans[c].vel.x=-chans[c].vel.x;
			/* Update channel position */
			chans[c].pos.z+=chans[c].vel.z*TIMERPERIOD/1000;
			chans[c].pos.x+=chans[c].vel.x*TIMERPERIOD/1000;
			BASS_ChannelSet3DPosition(chans[c].channel,&chans[c].pos,NULL,&chans[c].vel);
		}
		/* Draw the channel position indicator */
		x=cx+(int)((cx-7)*chans[c].pos.x/MAXDIST);
		y=cy-(int)((cy-7)*chans[c].pos.z/MAXDIST);
		if (chan==c)
			SelectObject(dc,brush1);
		else
			SelectObject(dc,brush2);
		Ellipse(dc,x-6,y-6,x+6,y+6);
	}
	/* Apply the 3D changes */
	BASS_Apply3D();

	RestoreDC(dc,save);
	DeleteObject(pen);
	ReleaseDC(win,dc);
}

/* Update the button states */
void UpdateButtons()
{
	int a;
	EnableWindow(ITEM(ID_REMOVE),chan==-1?FALSE:TRUE);
	EnableWindow(ITEM(ID_PLAY),chan==-1?FALSE:TRUE);
	EnableWindow(ITEM(ID_STOP),chan==-1?FALSE:TRUE);
	for (a=0;a<5;a++)
		EnableWindow(ITEM(ID_LEFT+a),chan==-1?FALSE:TRUE);
	if (chan!=-1) {
		MESS(ID_LEFT,BM_SETCHECK,chans[chan].dir==ID_LEFT?BST_CHECKED:BST_UNCHECKED,0);
		MESS(ID_RIGHT,BM_SETCHECK,chans[chan].dir==ID_RIGHT?BST_CHECKED:BST_UNCHECKED,0);
		MESS(ID_FRONT,BM_SETCHECK,chans[chan].dir==ID_FRONT?BST_CHECKED:BST_UNCHECKED,0);
		MESS(ID_BEHIND,BM_SETCHECK,chans[chan].dir==ID_BEHIND?BST_CHECKED:BST_UNCHECKED,0);
		MESS(ID_NONE,BM_SETCHECK,!chans[chan].dir?BST_CHECKED:BST_UNCHECKED,0);
	}
}

BOOL CALLBACK dialogproc(HWND h,UINT m,WPARAM w,LPARAM l)
{
	static OPENFILENAME ofn;
	static char path[MAX_PATH];

	switch (m) {
		case WM_COMMAND:
			switch (LOWORD(w)) {
				case ID_EAX:
					/* Change the EAX environment */
					if (HIWORD(w)==CBN_SELCHANGE) {
						int s=EM(CB_GETCURSEL,0,0);
						if (!s)
							BASS_SetEAXParameters(-1,0,-1,-1); // off (volume=0)
						else
							BASS_SetEAXParameters(s-1,-1,-1,-1);
					}
					break;
				case ID_LIST:
					/* Change the selected channel */
					if (HIWORD(w)!=LBN_SELCHANGE) break;
					chan=LM(LB_GETCURSEL,0,0);
					if (chan==LB_ERR) chan=-1;
					UpdateButtons();
					break;
				case ID_LEFT:
					chans[chan].dir=ID_LEFT;
					/* Make the channel move past the left of you */
					/* Set speed in m/s */
					chans[chan].vel.z=SPEED;
					chans[chan].vel.x=0;
					/* Set positon to the left */
					chans[chan].pos.x=-6;
					/* Reset display */
					InvalidateRect(GetDlgItem(h,ID_DISPLAY),NULL,TRUE);
					break;
				case ID_RIGHT:
					chans[chan].dir=ID_RIGHT;
					/* Make the channel move past the right of you */
					chans[chan].vel.z=SPEED;
					chans[chan].vel.x=0;
					/* Set positon to the right */
					chans[chan].pos.x=6;
					InvalidateRect(GetDlgItem(h,ID_DISPLAY),NULL,TRUE);
					break;
				case ID_FRONT:
					chans[chan].dir=ID_FRONT;
					/* Make the channel move past the front of you */
					chans[chan].vel.x=SPEED;
					chans[chan].vel.z=0;
					/* Set positon to in front */
					chans[chan].pos.z=6;
					InvalidateRect(GetDlgItem(h,ID_DISPLAY),NULL,TRUE);
					break;
				case ID_BEHIND:
					chans[chan].dir=ID_BEHIND;
					/* Make the channel move past the back of you */
					chans[chan].vel.x=SPEED;
					chans[chan].vel.z=0;
					/* Set positon to behind */
					chans[chan].pos.z=-6;
					InvalidateRect(GetDlgItem(h,ID_DISPLAY),NULL,TRUE);
					break;
				case ID_NONE:
					chans[chan].dir=0;
					/* Make the channel stop moving */
					chans[chan].vel.x=chans[chan].vel.z=0;
					break;
				case ID_ADD:
					{
						char file[MAX_PATH]="";
						DWORD newchan;
						ofn.lpstrFile=file;
						if (GetOpenFileName(&ofn)) {
							memcpy(path,file,ofn.nFileOffset);
							path[ofn.nFileOffset-1]=0;
							/* Load a music or sample from "file" */
							if ((newchan=BASS_MusicLoad(FALSE,file,0,0,BASS_MUSIC_RAMP|BASS_MUSIC_LOOP|BASS_MUSIC_3D,0))
								|| (newchan=BASS_SampleLoad(FALSE,file,0,0,1,BASS_SAMPLE_LOOP|BASS_SAMPLE_3D|BASS_SAMPLE_MONO))) {
								Channel *c;
								chanc++;
								chans=(Channel*)realloc((void*)chans,chanc*sizeof(Channel));
								c=chans+chanc-1;
								memset(c,0,sizeof(Channel));
								c->channel=newchan;
								BASS_SampleGetChannel(newchan,FALSE); // initialize sample channel
								LM(LB_ADDSTRING,0,file);
							} else
								Error("Can't load file (note samples must be mono)");
						}
					}
					break;
				case ID_REMOVE:
					{
						Channel *c=chans+chan;
						BASS_SampleFree(c->channel);
						BASS_MusicFree(c->channel);
						memcpy(c,c+1,(chanc-chan-1)*sizeof(Channel));
						chanc--;
						LM(LB_DELETESTRING,chan,0);
						chan=-1;
						UpdateButtons();
						InvalidateRect(GetDlgItem(h,ID_DISPLAY),NULL,TRUE);
					}
					break;
				case ID_PLAY:
					BASS_ChannelPlay(chans[chan].channel,FALSE);
					break;
				case ID_STOP:
					BASS_ChannelPause(chans[chan].channel);
					break;
				case IDCANCEL:
					DestroyWindow(h);
					break;
			}
			break;

		case WM_HSCROLL:
			if (l) {
				int pos=SendMessage((HWND)l,TBM_GETPOS,0,0);
				switch (GetDlgCtrlID((HWND)l)) {
					case ID_ROLLOFF: // change the rolloff factor
						BASS_Set3DFactors(-1,pow(2,(pos-10)/5.0),-1);
						break;
					case ID_DOPPLER: // change the doppler factor
						BASS_Set3DFactors(-1,-1,pow(2,(pos-10)/5.0));
						break;
				}
			}
			break;

		case WM_INITDIALOG:
			win=h;
			brush1=CreateSolidBrush(0xff);
			brush2=CreateSolidBrush(0);

			EM(CB_ADDSTRING,0,"Off");
			EM(CB_ADDSTRING,0,"Generic");
			EM(CB_ADDSTRING,0,"Padded Cell");
			EM(CB_ADDSTRING,0,"Room");
			EM(CB_ADDSTRING,0,"Bathroom");
			EM(CB_ADDSTRING,0,"Living Room");
			EM(CB_ADDSTRING,0,"Stone Room");
			EM(CB_ADDSTRING,0,"Auditorium");
			EM(CB_ADDSTRING,0,"Concert Hall");
			EM(CB_ADDSTRING,0,"Cave");
			EM(CB_ADDSTRING,0,"Arena");
			EM(CB_ADDSTRING,0,"Hangar");
			EM(CB_ADDSTRING,0,"Carpeted Hallway");
			EM(CB_ADDSTRING,0,"Hallway");
			EM(CB_ADDSTRING,0,"Stone Corridor");
			EM(CB_ADDSTRING,0,"Alley");
			EM(CB_ADDSTRING,0,"Forest");
			EM(CB_ADDSTRING,0,"City");
			EM(CB_ADDSTRING,0,"Mountains");
			EM(CB_ADDSTRING,0,"Quarry");
			EM(CB_ADDSTRING,0,"Plain");
			EM(CB_ADDSTRING,0,"Parking Lot");
			EM(CB_ADDSTRING,0,"Sewer Pipe");
			EM(CB_ADDSTRING,0,"Under Water");
			EM(CB_ADDSTRING,0,"Drugged");
			EM(CB_ADDSTRING,0,"Dizzy");
			EM(CB_ADDSTRING,0,"Psychotic");
			EM(CB_SETCURSEL,0,0);

			MESS(ID_ROLLOFF,TBM_SETRANGE,FALSE,MAKELONG(0,20));
			MESS(ID_ROLLOFF,TBM_SETPOS,TRUE,10);
			MESS(ID_DOPPLER,TBM_SETRANGE,FALSE,MAKELONG(0,20));
			MESS(ID_DOPPLER,TBM_SETPOS,TRUE,10);

			SetTimer(h,1,TIMERPERIOD,&Update);
			GetCurrentDirectory(MAX_PATH,path);
			memset(&ofn,0,sizeof(ofn));
			ofn.lStructSize=sizeof(ofn);
			ofn.hwndOwner=h;
			ofn.nMaxFile=MAX_PATH;
			ofn.lpstrInitialDir=path;
			ofn.Flags=OFN_HIDEREADONLY|OFN_EXPLORER;
			ofn.lpstrFilter="wav/aif/mo3/xm/mod/s3m/it/mtm/umx\0*.wav;*.aif;*.mo3;*.xm;*.mod;*.s3m;*.it;*.mtm;*.umx\0"
				"All files\0*.*\0\0";
			return 1;

		case WM_DESTROY:
			KillTimer(h,1);
			DeleteObject(brush1);
			DeleteObject(brush2);
			if (chans) free(chans);
			PostQuitMessage(0);
			break;
	}
	return 0;
}


// Simple device selector dialog stuff begins here
BOOL CALLBACK devicedialogproc(HWND h,UINT m,WPARAM w,LPARAM l)
{
	switch (m) {
		case WM_COMMAND:
			switch (LOWORD(w)) {
				case ID_DEVLIST:
					if (HIWORD(w)!=LBN_DBLCLK) break;
				case IDOK:
					{
						int device=SendDlgItemMessage(h,ID_DEVLIST,LB_GETCURSEL,0,0);
						device=SendDlgItemMessage(h,ID_DEVLIST,LB_GETITEMDATA,device,0); // get device #
						EndDialog(h,device);
					}
					break;
			}
			break;

		case WM_INITDIALOG:
			{
				char text[100],*d;
				int c,idx;
				for (c=1;d=BASS_GetDeviceDescription(c);c++) { // device 1 = 1st real device
					strcpy(text,d);
					/* Check if the device supports 3D */
					if (!BASS_Init(c,44100,BASS_DEVICE_3D,h,NULL))
						continue; // no 3D support
					if (BASS_GetEAXParameters(NULL,NULL,NULL,NULL))
						strcat(text," [EAX]"); // it has EAX
					BASS_Free();
					idx=SendDlgItemMessage(h,ID_DEVLIST,LB_ADDSTRING,0,(int)text);
					SendDlgItemMessage(h,ID_DEVLIST,LB_SETITEMDATA,idx,c); // store device #
					
				}
				SendDlgItemMessage(h,ID_DEVLIST,LB_SETCURSEL,0,0);
			}
			return 1;
	}
	return 0;
}
// Device selector stuff ends here

int PASCAL WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance,LPSTR lpCmdLine, int nCmdShow)
{
	MSG msg;
	int device;

	/* Check that BASS 2.2 was loaded */
	if (BASS_GetVersion()!=MAKELONG(2,2)) {
		MessageBox(0,"BASS version 2.2 was not loaded","Incorrect BASS.DLL",0);
		return 0;
	}

	{ // enable trackbar support
		INITCOMMONCONTROLSEX cc={sizeof(cc),ICC_BAR_CLASSES};
		InitCommonControlsEx(&cc);
	}

	/* Create the main window */
	if (!CreateDialog(hInstance,MAKEINTRESOURCE(1000),NULL,&dialogproc)) {
		Error("Can't create window");
		return 0;
	}

	/* Let the user choose an output device */
	device=DialogBox(hInstance,(char*)2000,win,&devicedialogproc);

	/* Initialize the output device with 3D support */
	if (!BASS_Init(device,44100,BASS_DEVICE_3D,win,NULL)) {
		Error("Can't initialize output device");
		EndDialog(win,0);
		return 0;
	}

	/* Use meters as distance unit, real world rolloff, real doppler effect */
	BASS_Set3DFactors(1,1,1);
	/* Turn EAX off (volume=0), if error then EAX is not supported */
	if (BASS_SetEAXParameters(-1,0,-1,-1))
		EnableWindow(GetDlgItem(win,ID_EAX),TRUE);

	while (GetMessage(&msg,NULL,0,0)>0) {
		if (!IsDialogMessage(win,&msg)) {
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	
	BASS_Free();

	return 0;
}
