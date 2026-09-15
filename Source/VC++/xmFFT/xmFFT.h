
// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the XMFFT_EXPORTS
// symbol defined on the command line. this symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// XMFFT_API functions as being imported from a DLL, wheras this DLL sees symbols
// defined with this macro as being exported.
#ifdef XMFFT_EXPORTS
#define XMFFT_API __declspec(dllexport) WINAPI
#else
#define XMFFT_API __declspec(dllimport)
#endif

#include <windef.h>

void prepareMDCT (float *data,long Max,int Barre);
void GetResult(float *pfImagBuffer, float *pfRealBuffer,int iNum);
