# Microsoft Developer Studio Project File - Name="mp3dec" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Static Library" 0x0104

CFG=mp3dec - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "mp3dec.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "mp3dec.mak" CFG="mp3dec - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "mp3dec - Win32 Release" (based on "Win32 (x86) Static Library")
!MESSAGE "mp3dec - Win32 Debug" (based on "Win32 (x86) Static Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""$/mp3dec", LWAAAAAA"
# PROP Scc_LocalPath "."
CPP=cl.exe
RSC=rc.exe

!IF  "$(CFG)" == "mp3dec - Win32 Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Release"
# PROP Intermediate_Dir "Release"
# PROP Target_Dir ""
# ADD BASE CPP /nologo /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_MBCS" /D "_LIB" /YX /FD /c
# ADD CPP /nologo /G6 /MD /W3 /GX /O2 /I ".\include" /D "WIN32" /D "NDEBUG" /D "_MBCS" /D "_LIB" /FR /YX /FD /c
# ADD BASE RSC /l 0x411 /d "NDEBUG"
# ADD RSC /l 0x409 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LIB32=link.exe -lib
# ADD BASE LIB32 /nologo
# ADD LIB32 /nologo

!ELSEIF  "$(CFG)" == "mp3dec - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Debug"
# PROP Intermediate_Dir "Debug"
# PROP Target_Dir ""
# ADD BASE CPP /nologo /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_MBCS" /D "_LIB" /YX /FD /GZ /c
# ADD CPP /nologo /GB /MTd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_MBCS" /D "_LIB" /YX /FD /GZ /c
# ADD BASE RSC /l 0x411 /d "_DEBUG"
# ADD RSC /l 0x409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LIB32=link.exe -lib
# ADD BASE LIB32 /nologo
# ADD LIB32 /nologo

!ENDIF 

# Begin Target

# Name "mp3dec - Win32 Release"
# Name "mp3dec - Win32 Debug"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\src\bstream.c
# End Source File
# Begin Source File

SOURCE=.\src\fdct.c
# End Source File
# Begin Source File

SOURCE=.\src\imdct.c
# End Source File
# Begin Source File

SOURCE=.\src\l1dec.c
# End Source File
# Begin Source File

SOURCE=.\src\l1init.c
# End Source File
# Begin Source File

SOURCE=.\src\l2dec.c
# End Source File
# Begin Source File

SOURCE=.\src\l2init.c
# End Source File
# Begin Source File

SOURCE=.\src\l3alias.c
# End Source File
# Begin Source File

SOURCE=.\src\l3dec.c
# End Source File
# Begin Source File

SOURCE=.\src\l3huff.c
# End Source File
# Begin Source File

SOURCE=.\src\l3hybrid.c
# End Source File
# Begin Source File

SOURCE=.\src\l3init.c
# End Source File
# Begin Source File

SOURCE=.\src\l3msis.c
# End Source File
# Begin Source File

SOURCE=.\src\l3quant.c
# End Source File
# Begin Source File

SOURCE=.\src\l3sf.c
# End Source File
# Begin Source File

SOURCE=.\src\l3side.c
# End Source File
# Begin Source File

SOURCE=.\src\mp3dec.c
# End Source File
# Begin Source File

SOURCE=.\src\sbt.c
# End Source File
# Begin Source File

SOURCE=.\src\sbtb.c
# End Source File
# Begin Source File

SOURCE=.\src\window.c
# End Source File
# Begin Source File

SOURCE=.\src\windowb.c
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\include\bstream.h
# End Source File
# Begin Source File

SOURCE=.\include\l3huff.h
# End Source File
# Begin Source File

SOURCE=.\include\layer3.h
# End Source File
# Begin Source File

SOURCE=.\include\mp3dec.h
# End Source File
# Begin Source File

SOURCE=..\xmMP30159\include\MP3Output.h
# End Source File
# End Group
# End Target
# End Project
