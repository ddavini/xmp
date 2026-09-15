#include "stdafx.h"
#include "xmOGG.h"
#include "OggPlayer.h"

OggPlayer op;


BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
    switch (ul_reason_for_call)
	{
		case DLL_PROCESS_ATTACH:
		case DLL_THREAD_ATTACH:
		case DLL_THREAD_DETACH:
		case DLL_PROCESS_DETACH:
			break;
    }
    return TRUE;
}

int __stdcall InitOgg (HWND hwnd, int global){
	op.InitDirectSound(hwnd, global);
	return 0;
}

int __stdcall PlayOggFile (char * FileName){
	
	//op.InitDirectSound(hwnd);
	op.OpenOgg(FileName);
	op.Play();
	
	//while(op.IsPlaying());
	//op.Close();
	
	return 0;
}

int __stdcall OggIsPlaying (){
	return op.IsPlaying();
}

int __stdcall StopOgg (){
	op.Stop();	
	return 0;
}

int __stdcall UpdateOgg () {
	op.Update();
	return 0;
}

int __stdcall CloseOgg () {
	op.~OggPlayer();
	op.Close();
	return 0;
}

int __stdcall TerminateOgg () {
	op.Close();
	TerminateProcess(GetCurrentProcess,0); 
	return 0;
}

int __stdcall SetOggVolume (int Volume) {
	op.SetVolume(Volume);
	return 0;
}

int __stdcall SetOggPan (int Position) {
	op.SetPan(Position);
	return 0;
}