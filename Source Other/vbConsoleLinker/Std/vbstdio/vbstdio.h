
// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the VBSTDIO_EXPORTS
// symbol defined on the command line. this symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// VBSTDIO_API functions as being imported from a DLL, wheras this DLL sees symbols
// defined with this macro as being exported.

#define PERR(bSuccess, api){if(!(bSuccess)) printf("%s:Error %d from %s \
on line %d\n", __FILE__, GetLastError(), api, __LINE__);}

#ifdef VBSTDIO_EXPORTS
#define VBSTDIO_API __declspec(dllexport)
#else
#define VBSTDIO_API __declspec(dllimport)
#endif

extern VBSTDIO_API int nVbstdio;

VBSTDIO_API int __stdcall vbprintf(const char* Str);
VBSTDIO_API void __stdcall vbclrscr();
VBSTDIO_API void __stdcall vbsetcmdcursorposition(COORD Coordinates);