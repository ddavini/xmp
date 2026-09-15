// xmMP3.cpp : DLL アプリケーション用のエントリ ポイントを定義します。
//

#include "..\\include\\stdafx.h"
#include "..\\include\\xmMP3.h"
#include "string.h"



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

xmMP3_API BOOL __stdcall xmMP3_getWaveData(WAVE_DATA* pWaveData)
{
	//return input_getWaveData(pWaveData);
	return FALSE;
}

xmMP3_API void __stdcall xmMP3_getSpectrum(int* pSpecL, int* pSpecR)
{
	input_getSpectrum(pSpecL, pSpecR);
	return;
}

xmMP3_API void __stdcall xmMP3_getWave(int* pWaveL, int* pWaveR)
{
	input_getWave(pWaveL, pWaveR);
}

//コールバック関数
xmMP3_API BOOL __stdcall xmMP3_callback(INPUT_MSG_PROC pProc)
{
	return input_setCallback(pProc);
}

xmMP3_API BOOL __stdcall xmMP3_startCallback()
{
	return input_startCallback();
}

xmMP3_API BOOL __stdcall xmMP3_stopCallback()
{
	return input_stopCallback();
}

xmMP3_API BOOL __stdcall xmMP3_playDecodeWave(const char* pszWaveName)
{
	return input_playDecodeWave(pszWaveName);
}

xmMP3_API BOOL __stdcall xmMP3_decodeWave(const char* pszWaveName)
{
	return input_decodeWave(pszWaveName);
}

xmMP3_API BOOL __stdcall xmMP3_getMpegInfo(MPEG_INFO* pMpegInfo)
{
	return getMpegInfo(pMpegInfo);
}

xmMP3_API long __stdcall xmMP3_getWinampPlayMs()
{
	return input_getWinampPlayMs();
}

xmMP3_API int __stdcall xmMP3_getWinampTotalSec()
{
	return input_getWinampTotalSec();
}

xmMP3_API int __stdcall xmMP3_getPlayBitRate()
{
	return input_getPlayBitRate();
}

xmMP3_API int __stdcall xmMP3_getLastErrorNo()
{
	return input_getLastErrorNo();
}

// Obsolete - Retro Compatibilty Only
xmMP3_API int __stdcall xmMP3_getVersion()
{
	int Version = 9999;
	return Version;
}

xmMP3_API BOOL __stdcall xmMP3_setxmMP3Option(xmMP3_OPTION* pxmMP3Option)
{
	return input_SetxmMP3Option(pxmMP3Option);
}

xmMP3_API void __stdcall xmMP3_getxmMP3Option(xmMP3_OPTION* pxmMP3Option)
{
	input_GetxmMP3Option(pxmMP3Option);
}

xmMP3_API BOOL __stdcall xmMP3_setDecodeOption(DEC_OPTION* pDecOption)
{
	return input_SetDecodeOption(pDecOption);
}

xmMP3_API void __stdcall xmMP3_getDecodeOption(DEC_OPTION* pDecOption)
{
	input_GetDecodeOption(pDecOption);
}

xmMP3_API long __stdcall xmMP3_debug()
{
	return input_debug();
}

xmMP3_API void __stdcall xmMP3_setEqualizer(int* pTable)
{
	input_setEqualizer(pTable);
}

xmMP3_API void __stdcall xmMP3_setFftWindow(int window)
{
	input_setFftWindow(window);
}

xmMP3_API void __stdcall xmMP3_setWaveOutDeviceId(int id){
	input_setWaveOutDeviceId(id);
}

//基本操作系
xmMP3_API BOOL __stdcall xmMP3_setStepPitch(int pitch, int frames)
{
	return input_setStepPitch(pitch, frames);
}

xmMP3_API int __stdcall xmMP3_getStepPitch()
{
	return input_getStepPitch();
}

xmMP3_API BOOL __stdcall xmMP3_reload()
{
	return input_reload();
}

xmMP3_API BOOL __stdcall xmMP3_setPitch(int pitch)
{
	return input_setPitch(pitch);
}

xmMP3_API int __stdcall xmMP3_getPitch()
{
	return input_getPitch();
}

xmMP3_API BOOL __stdcall xmMP3_init(int flag)
{
	return initDec();
}


xmMP3_API BOOL __stdcall xmMP3_open(const char* pszName, InputInfo* pInfo)
{

	return input_open(pszName, pInfo);
	
}

xmMP3_API BOOL __stdcall xmMP3_close()
{
	return input_close();

}

xmMP3_API int __stdcall xmMP3_getState(int* sec)
{

	*sec = getTime();

	return input_getState();

}


xmMP3_API BOOL __stdcall xmMP3_play()
{
	return input_play();
}

//X-MaD 28/08/01
xmMP3_API BOOL __stdcall xmMP3_SetBufferSize(int BufferRedimSize)
{
	return output_SetBufferRedim(BufferRedimSize);
}
//X-MaD 28/08/01

//X-MaD 26/03/02
xmMP3_API BOOL __stdcall xmMP3_SetReadHaedA(int NEqualFrame, int MaxHeadSize = 102400)
{
	return input_SetReadHaedA(NEqualFrame,MaxHeadSize);
}
//X-MaD 26/03/02

xmMP3_API void __stdcall xmMP3_setXSoundMode(int mode = 0,int option = 0)
{
	output_SetXSound(mode,option);
}

xmMP3_API double __stdcall xmMP3_getXNormLevel()
{
	return output_getXNormLevel();
}

xmMP3_API double __stdcall xmMP3_setXNormLevel(double level)
{
	return output_setXNormLevel(level);
}


xmMP3_API int __stdcall xmMP3_getXSoundMode()
{
	return output_GetXSound();
}


xmMP3_API BOOL __stdcall xmMP3_stop()
{
	return input_stop();
}

xmMP3_API BOOL __stdcall xmMP3_pause()
{
	return input_pause();
}

xmMP3_API BOOL __stdcall xmMP3_restart()
{
	return input_restart();
}

xmMP3_API BOOL __stdcall xmMP3_seek(int seekSec)
{
	input_seekBySec(seekSec);
	return TRUE;
}

xmMP3_API BOOL __stdcall xmMP3_setVolume(int left, int right)
{
	return input_setVolume(left, right);
}

xmMP3_API BOOL __stdcall xmMP3_getVolume(int* left, int* right)
{
	return input_getVolume(left, right);
}

xmMP3_API int __stdcall xmMP3_getWaveOutSupport()
{
	return input_getWaveOutSupport();
}

xmMP3_API void __stdcall xmMP3_setSoftVolume(int left, int right)
{
	input_setSoftVolume(left, right);
}

xmMP3_API void __stdcall xmMP3_getSoftVolume(int* left, int* right)
{
	input_getSoftVolume(left, right);
}

xmMP3_API int __stdcall xmMP3_getPlayFlames()
{
	return input_getPlayFlames();
}

xmMP3_API BOOL __stdcall xmMP3_setPlayFlames(int flames)
{
	input_seekByFrame(flames);
	return TRUE;
}

xmMP3_API long __stdcall xmMP3_getPlaySamples()
{
	return input_getPlaySamples();
}

xmMP3_API int __stdcall xmMP3_getTotalSamples()
{
	return input_getTotalSamples();
}

xmMP3_API BOOL __stdcall xmMP3_setPlaySamples(int sample)
{
	input_seekBySample(sample);
	return TRUE;
}

xmMP3_API void __stdcall xmMP3_setFadeIn(int on)
{
	input_setFadeIn(on);
}

xmMP3_API void __stdcall xmMP3_setFadeOut(int on)
{
	input_setFadeOut(on);
}

xmMP3_API void __stdcall xmMP3_fadeOut()
{
	input_fadeOut();
}

xmMP3_API void __stdcall xmMP3_setOverTime(int on)
{
	input_setOverTime(on);
}