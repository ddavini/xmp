#include "..\\include\MP3Input.h"

typedef unsigned int (__stdcall *THREADFUNC)(void* pParam);

void __stdcall _NullProc(int msg, int param) {}

void input_setWaveOutDeviceId(int id){
	output_setWaveOutDeviceId(id);
}

void input_setFftWindow(int window)
{
	output_setFftWindow(window);
}

int input_GetState() {

	return m_state;
}

void input_setOverTime(int on)
{
	if ( on == 0 ) {
		m_timeFlag = FALSE;
	} else {
		m_timeFlag = TRUE;
	}
}

int input_getLastErrorNo()
{
	return m_errNo;
}

void input_setErrNo(int no)
{
	m_errNo = no;
}

void input_fadeOut()
{
	output_getVolume(&m_olvol,&m_orvol);
	m_lvol = m_olvol;
	m_rvol = m_orvol;
	m_fadeOutFlag = true;
}

void input_fadeOutVolume()
{
	if ( m_fadeOutFlag ){
		if ( m_lvol > 0 )m_lvol--;
		if ( m_rvol > 0 )m_rvol--;
		output_setVolume(m_lvol, m_rvol);

		if (m_lvol == 0 && m_rvol == 0) {
			m_lvol=m_olvol;
			m_rvol=m_orvol;
			m_fadeOutFlag = false;
			output_stop();
		}
	}
}

int input_getDecodeOutputSize()
{
	return (int)output_sampleToBytes(m_firstSample);
}

int input_getTotalOutputSize()
{
	return (int)output_sampleToBytes(input_getTotalSamples());
}


void input_setFadeIn(int on)
{
	output_setFadeIn(on);
}

void input_setFadeOut(int on)
{
	output_setFadeOut(on);
}


BOOL input_reload()
{
	if ( m_state == STATE_STOP ) {
		m_errNo = ERR_STATE_STOP;
		return FALSE;
	}

	int w_state = m_state;
	DWORD curSample = output_getCurrentSample() + m_firstSample;

	input_stop();
	input_play();
	input_seekBySample(curSample);
	if ( w_state == STATE_PAUSE ) 
		input_pause();
	return TRUE;
}

BOOL input_setStepPitch(int pitch, int frames)
{
	//WinNT ébíËëŒçÙ
	if ( m_platForm == 1 ) {
		m_errNo = ERR_PLATFORM_NT;
		return FALSE;
	}

	if ( pitch < 1 || frames < 1)
	{
		m_errNo = ERR_INVALID_VALUE;
		return FALSE;
	}
	int w_state = m_state;
	DWORD curSample = output_getCurrentSample() + m_firstSample;

	if ( pitch == 1 ) {
		m_outputBlock_in = m_outputBlock_in_tmp;
		m_stepFlag = FALSE;
	} else {
		if ( m_stepFlag == FALSE ) m_outputBlock_in_tmp = m_outputBlock_in; 
		m_outputBlock_in = frames;
		m_stepFlag = TRUE;
	}

	output_setStepPitch(pitch);
	if ( m_state != STATE_STOP) {
		input_stop();
		input_play();
		input_seekBySample(curSample);
		if ( w_state == STATE_PAUSE )
			input_pause();
	}
	return TRUE;
}

int input_getStepPitch()
{
	return output_getStepPitch();
}

BOOL input_setPitch(int pitch)
{
	if ( pitch < 20 )
	{
		m_errNo = ERR_INVALID_VALUE;
		return FALSE;
	}
	output_setPitch(pitch);
	int w_state = m_state;
	DWORD curSample = output_getCurrentSample() + m_firstSample;
	if ( m_state != STATE_STOP) {
		input_stop();
		input_play();
		input_seekBySample(curSample);
		if ( w_state == STATE_PAUSE )
			input_pause();
	}
	return TRUE;
}

int input_getPitch()
{
	return output_getPitch();
}

void input_callBackTime()
{
	if ( callBackFlag == TRUE ) {
		m_playSamples = output_getCurrentSample();
		m_playSec = (int)(input_sampleToSec(m_playSamples + m_firstSample) * m_playSpeed);
		if ( m_timeFlag && m_playSec > m_totalSec ) m_playSec = m_totalSec;
		m_pProc(INPUT_MSG_PLAY, m_playSec);
	}
}

long input_getWinampPlayMs()
{
	long winampMs;
	if ( m_decInfo.frequency == 0 || m_totalSec == 0 ) {
		return 0;
	}

	//î{ë¨ââëtëŒâû
	double playMs;
	DWORD playSamples = output_getCurrentSample();
	playMs = (((float)playSamples + (float)m_firstSample) * (float)1000 / (float)m_decInfo.frequency) / (float)1000;
	playMs = playMs * m_playSpeed;
	playMs = playMs * (float)((float)m_winampTotalSec/(float)m_totalSec);
	winampMs = (long)(playMs * 1000);
	if (winampMs > (m_winampTotalSec * 1000) ) winampMs = m_winampTotalSec * 1000;
	return winampMs;
}

int input_getWinampTotalSec()
{
	return m_winampTotalSec;
}

int input_getPlayBitRate()
{
	if ( m_waveFlag == FALSE ) {
		return m_bitRate / 1000;
	} else {
		return m_bitRate;
	}
}

int input_getPlayFlames()
{
	m_playFrames = 0;

	//ÉtÉåÅ[ÉÄêîéÊìæ
	if ( m_decInfo.outputSize > 0 ){
		if ( decFlag == 1 ){	
			m_playFrames = (int)(((m_firstSample + input_bytesToSample(decodePos)) * m_decInfo.bitsPerSample * m_decInfo.channels / 8) / m_decInfo.outputSize);
		} else {
			m_playFrames = (int)(output_sampleToBytes(output_getCurrentSample() + m_firstSample) / m_decInfo.outputSize);
		}
	}
	return m_playFrames;
}

DWORD input_getPlaySamples()
{
	if ( decFlag == 1 ) {
		return (m_firstSample + input_bytesToSample(decodePos));
	}
	return output_getCurrentSample() + m_firstSample;
}

int input_getTotalSamples()
{

	return input_bytesToSample(m_decInfo.frames * m_decInfo.outputSize);
	
}

long input_debug()
{
	return debugNo;
}

void input_setDebug(long no)
{
	debugNo = no;
}

BOOL input_SetxmMP3Option(xmMP3_OPTION* pxmMP3Option)
{

	if (m_state != STATE_STOP) {
		//í‚é~íÜÇÃÇ›èàóùÇ∑ÇÈ
		m_errNo = ERR_NOT_STATE_STOP;
		return FALSE;
	}

	//ÉIÉvÉVÉáÉìílÉ`ÉFÉbÉN
	if (pxmMP3Option->inputBlock <= 0) {
		m_errNo = ERR_INVALID_VALUE;
		return FALSE;
	}

	if (pxmMP3Option->outputBlock <= 0) {
		m_errNo = ERR_INVALID_VALUE;
		return FALSE;
	}

	if (pxmMP3Option->inputSleep < 0) {
		m_errNo = ERR_INVALID_VALUE;
		return FALSE;
	}

	if (pxmMP3Option->outputSleep < 0) {
		m_errNo = ERR_INVALID_VALUE;
		return FALSE;
	}

	m_inputBlock	= pxmMP3Option->inputBlock;
	m_outputBlock_in	= pxmMP3Option->outputBlock;
	m_outputBlock_in_tmp = m_outputBlock_in;
	m_inputSleep	= pxmMP3Option->inputSleep;
	//m_outputSleep	= pxmMP3Option->outputSleep;

	output_SetxmMP3Option(pxmMP3Option);

	return TRUE;

}

void input_GetxmMP3Option(xmMP3_OPTION* pxmMP3Option){
	pxmMP3Option->inputBlock = m_inputBlock;
	pxmMP3Option->outputBlock = m_outputBlock_in;
	pxmMP3Option->inputSleep = m_inputSleep;
	//pxmMP3Option->outputSleep = m_outputSleep;
	output_GetxmMP3Option(pxmMP3Option);
	return;
}

BOOL input_SetDecodeOption(DEC_OPTION* pDecOption)
{
	double curRed;
	double Red;

	if (m_state != STATE_STOP) {
		//í‚é~íÜÇÃÇ›èàóùÇ∑ÇÈ
		m_errNo = ERR_NOT_STATE_STOP;
		return FALSE;
	}

	//ÉIÉvÉVÉáÉìílÉ`ÉFÉbÉN
	if (pDecOption->reduction != 0 &&
	    pDecOption->reduction != 1 &&
		pDecOption->reduction != 2) {
		m_errNo = ERR_INVALID_VALUE;
		return FALSE;
	}

	if (pDecOption->convert != 0 &&
	    pDecOption->convert != 1) {
		m_errNo = ERR_INVALID_VALUE;
		return FALSE;
	}

	if (pDecOption->freqLimit <= 0) {
		m_errNo = ERR_INVALID_VALUE;
		return FALSE;
	}

	decodeOption.reduction = pDecOption->reduction;
	decodeOption.convert = pDecOption->convert;
	decodeOption.freqLimit = pDecOption->freqLimit;

	if ( pDecOption->reduction == 0 ) {
		curRed = 4;
	} else if ( pDecOption->reduction == 1 ) {
		curRed = 2;
	} else {
		curRed = 1;
	}

	if ( m_reduction == 0 ) {
		Red = 4;
	} else if ( m_reduction == 1 ) {
		Red = 2;
	} else {
		Red = 1;
	}
	m_playSpeed = Red / curRed;

	return ::mp3SetDecodeOption(&decodeOption);
}

void input_GetDecodeOption(DEC_OPTION* pDecOption)
{
	mp3GetDecodeOption(&decodeOption);

	pDecOption->reduction = decodeOption.reduction;
	pDecOption->convert = decodeOption.convert;
	pDecOption->freqLimit = decodeOption.freqLimit;

	return;
}

int getTime()
{
	if (m_decInfo.frequency == 0 || m_decInfo.frequency == 0) return 0;

	if ( decFlag == 1 ){
		return input_sampleToSec(m_firstSample + input_bytesToSample(decodePos));
	}

	double playMs;
	DWORD playSamples = output_getCurrentSample();
	//î{ë¨ââëtëŒâû
	playMs = (((float)playSamples + (float)m_firstSample) * (float)1000 / (float)m_decInfo.frequency) / (float)1000;
	playMs = playMs * m_playSpeed;
	m_playSec = (int)playMs;
	if (m_timeFlag && m_playSec > m_totalSec) m_playSec = m_totalSec;
	return m_playSec;
}

BOOL getMpegInfo(MPEG_INFO* pMpegInfo)
{
	int				padding;

	//ÉtÉ@ÉCÉãñ¢ÉIÅ[ÉvÉì
	if (m_hFile == INVALID_HANDLE_VALUE) {
		m_errNo = ERR_MP3_FILE_NOT_OPEN;
		return FALSE;
	}

	pMpegInfo->version = m_decInfo.header.version;
	pMpegInfo->layer = m_decInfo.header.layer;
	pMpegInfo->crcDisable = (m_decInfo.header.error_prot) ? 0 : 1;
	pMpegInfo->mode = m_decInfo.header.mode;
	pMpegInfo->extension = m_decInfo.header.mode_ext;
	pMpegInfo->copyright = m_decInfo.header.copyright;
	pMpegInfo->original = m_decInfo.header.original;
	pMpegInfo->emphasis = m_decInfo.header.emphasis;
	pMpegInfo->channels = m_decInfo.channels;
	pMpegInfo->bitRate = m_decInfo.bitRate / 1000;
	pMpegInfo->fileSize = fileSize;
	pMpegInfo->samplingRate = m_decInfo.frequency;
	padding = m_decInfo.header.padding;

	if (m_decInfo.frames == 0) {
		m_decInfo.frames = dataSize / m_decInfo.minInputSize;
	}
	pMpegInfo->flames = m_decInfo.frames;
    
	int samples = input_bytesToSample(m_decInfo.frames * m_decInfo.outputSize);
	pMpegInfo->totalSec = input_sampleToSec(samples);

	return TRUE;
}


BOOL initDec()
{

	::mp3DecodeInit();

	OSVERSIONINFO	obsInfo;
	obsInfo.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
	::GetVersionEx(&obsInfo);
	if ( obsInfo.dwPlatformId == VER_PLATFORM_WIN32_NT ) {
		m_platForm = 1;
	} else {
		m_platForm = 0;
	}
	
	m_pProc			= _NullProc;
	m_hFile			= INVALID_HANDLE_VALUE;

	m_hThread		= NULL;
	m_threadID		= 0;
	m_state			= STATE_STOP;
	m_inputBuf		= NULL;
	m_outputBuf_in	= NULL;

	callBackFlag	= FALSE;

	debugNo = 0;

	initOutput();
	//initEnc();
	mp3SetEqualizer(NULL);
	return TRUE;
}

BOOL freeDec()
{
	::Sleep(50);
	callBackFlag = FALSE;
	m_pProc = _NullProc;

	if ( m_playFlag ) input_stop();
	input_close();

	if ( m_playFlag ) output_stop();
	if ( m_playFlag ) output_close();
	freeOutput();
	//freeEnc();

	::free(m_inputBuf);
	::free(m_outputBuf_in);
	::Sleep(50);
	return TRUE;
}

BOOL input_open(const char* pszName, InputInfo* pInfo)
{
	Sleep(10);
	
	if (pszName[0] == 0) {
	
		m_errNo = ERR_MP3_FILE_OPEN;
		return FALSE;
	}

	
	::input_close();
	m_waveFlag= FALSE;

	
	m_hFile = ::CreateFile(pszName, GENERIC_READ, FILE_SHARE_READ + FILE_SHARE_WRITE,
			NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (m_hFile == INVALID_HANDLE_VALUE) {
		m_errNo = ERR_MP3_FILE_OPEN;
		return FALSE;
	}

	m_openFlag = TRUE;

	m_firstSync = 0;
	dataSize = ::GetFileSize(m_hFile, NULL);
	fileSize = dataSize;

	::memset(pInfo, 0, sizeof(InputInfo));

	::readRIFF(m_hFile, &m_firstSync, &dataSize);
	if (RiffStatus == -1) {
		//MPEG,WAVE
		::input_close();
		m_errNo = ERR_UNKNOWN_FILE;
		return FALSE;
	} else if (RiffStatus == 1) {
		//WAVE
		m_decInfo.bitsPerSample = m_waveFormat.wBitsPerSample;
		m_decInfo.channels = m_waveFormat.nChannels;
		m_decInfo.dataSize = dataSize;
		m_decInfo.frequency = m_waveFormat.nSamplesPerSec;
		m_decInfo.bitRate = m_waveFormat.wBitsPerSample;		
		m_bitRate = m_decInfo.bitRate;							
		m_playSpeed = 1;
		m_firstSample = 0;
		m_decInfo.outputSize = DEFAULT_BUFFER_SIZE;
		if ( m_decInfo.bitRate == 8 ) m_decInfo.outputSize = m_decInfo.outputSize / 4;
		m_decInfo.frames = (int)(m_decInfo.dataSize / m_decInfo.outputSize);

		::strcpy(pInfo->szTrackName,  "WAVE_FORMAT_PCM");
		::strcpy(pInfo->szArtistName, "WAVE_FORMAT_PCM");
		pInfo->bitRate = m_waveFormat.wBitsPerSample;
		pInfo->channels = m_waveFormat.nChannels;
		pInfo->samplingRate = m_waveFormat.nSamplesPerSec;
		pInfo->totalSec = input_sampleToSec(input_bytesToSample(dataSize));
		m_totalSec = pInfo->totalSec;
		m_winampTotalSec = pInfo->totalSec;
		m_waveFlag= TRUE;
		return TRUE;
	}

//X-MaD

	BOOL	result;
	DWORD	bytesRead;
	long FileSize = GetFileSize(m_hFile, NULL);
	

	BYTE	*buf;
	int		sync;

	int InputBytes = 2048; //2 KB
	
	buf = (BYTE*)malloc(InputBytes);

	long NextBufRead = m_firstSync;
	int nvolte = 0;
	int Bitrate = 0;
	int Freq = 0;
	int firstSync = 0;
	bool primavolta = true;


	for (;;) {

		for (;;) {
			
//			Sleep(1);

			if (NextBufRead >= FileSize || NextBufRead >= m_MaxHeadSize
				|| firstSync >= (FileSize / 3)) {
				::input_close();
				return FALSE;
			}

			::SetFilePointer(m_hFile, NextBufRead, NULL, FILE_BEGIN);
			result = ::ReadFile(m_hFile, buf, InputBytes, &bytesRead, NULL);
			if (!result || bytesRead == 0) {
				m_errNo = ERR_MP3_FILE_READ;
				::input_close();
				return FALSE;
			}
			
			
			if (!::mp3FindSync(buf, InputBytes, &sync)) {
				//::input_close();
				m_errNo = ERR_FRAME_HEADER_NOT_FOUND;
				//return FALSE;
				NextBufRead = NextBufRead + InputBytes; //firstSync;
			} else {
				goto exitfor;
			}

		}

exitfor:

		firstSync = NextBufRead + sync;

		::SetFilePointer(m_hFile, firstSync, NULL, FILE_BEGIN);
		result = ::ReadFile(m_hFile, buf, InputBytes, &bytesRead, NULL);
		if (!result || bytesRead == 0) {
			::input_close();
			m_errNo = ERR_MP3_FILE_READ;
			return FALSE;
		}



		if (!::mp3GetDecodeInfo(buf, bytesRead, &m_decInfo, 1)) {
			NextBufRead = NextBufRead + 1; //firstSync;
		} 
		else {
			if (primavolta) {	
					NextBufRead = firstSync;
					m_firstSync = firstSync;
					primavolta = false;
			}
			if ((Bitrate == m_decInfo.bitRate && Freq == m_decInfo.frequency) || nvolte == 0) {
				if (firstSync == 0 || m_decInfo.bitRate == 0) {
					goto Exit;
				} else {
					nvolte = nvolte + 1;
				}
			} else {
				NextBufRead = firstSync;
				if (nvolte = 1) {
					primavolta = true;
				}
				nvolte = 0;
			}
			Bitrate = m_decInfo.bitRate;
			Freq = m_decInfo.frequency;
			NextBufRead = NextBufRead + 1;
			if (nvolte >= (m_NEqualFrame * 4)) {
				goto Exit;
			}
		}

	}

Exit:
//X-MaD 02/08/01

	m_firstSample = 0;

	pInfo->channels = m_decInfo.channels;
	pInfo->bitRate = m_decInfo.bitRate / 1000;
	m_bitRate = pInfo->bitRate;
	pInfo->samplingRate = m_decInfo.frequency;

	if (m_decInfo.frames == 0) {
		m_decInfo.frames = dataSize / m_decInfo.minInputSize;
	}

	int samples = input_bytesToSample(m_decInfo.frames * m_decInfo.outputSize);
	pInfo->totalSec = input_sampleToSec(samples);
	m_totalSec = pInfo->totalSec;
	
	//winamp
	double msPerFrame = ms_table[m_decInfo.header.layer-1][m_decInfo.header.fr_index];

	m_winampTotalSec = (m_decInfo.frames * (int)(msPerFrame) / 1000);

	mp3GetDecodeOption(&decodeOption);
	m_reduction = decodeOption.reduction;
	m_playSpeed = 1;

	return TRUE;
}


//X-MaD 26/03/02
BOOL input_SetReadHaedA(int NEqualFrame,int MaxHeadSize)
{
	if (NEqualFrame <= 1024 && NEqualFrame >= 0)
	{			
		m_NEqualFrame = NEqualFrame;
	} else {
		return FALSE;
	}
	m_MaxHeadSize = MaxHeadSize;
	if (MaxHeadSize >= 102400 || MaxHeadSize <= 0)
	{			
		m_MaxHeadSize = 999999;
		return TRUE;
	} else {
		return FALSE;
	}
}
//X-MaD 26/03/02

BOOL input_close()
{
	::Sleep(10);
	if (m_hFile == INVALID_HANDLE_VALUE) {
		m_errNo = ERR_MP3_FILE_NOT_OPEN;
		return FALSE;
	}

	if (m_state != STATE_STOP) {
		input_stop();
	}

	::CloseHandle(m_hFile);
	m_hFile = INVALID_HANDLE_VALUE;

	return TRUE;
}

int input_read(void* pData, int length)
{
	int bytesRead;
	if (!::ReadFile(m_hFile, pData, length, (DWORD*)&bytesRead, NULL)) {
		return 0;//exception
	}
	return bytesRead;
}

BOOL input_setCallback(INPUT_MSG_PROC pProc)
{
	if (m_state != STATE_STOP) {
		//í‚é~íÜà»äOÇÕê›íËïsâ¬
		m_errNo = ERR_NOT_STATE_STOP;
		return FALSE;
	}
	m_pProc = pProc;
	callBackFlag = TRUE;
	return TRUE;
}

BOOL input_startCallback()
{
	callBackFlag = TRUE;
	return TRUE;
}

BOOL input_stopCallback()
{
	callBackFlag = FALSE;
	return TRUE;
}

int input_getState()
{
	return m_state;
}

BOOL input_play()
{
	//Sleep(50);
	if (m_state != STATE_STOP || m_hThread != NULL ) {
		//í‚é~íÜà»äOÇÕé¿çsïsâ¬
		m_errNo = ERR_NOT_STATE_STOP;
		return FALSE;
	}

	decFlag = 0;
	m_fadeOutFlag = false;
	m_playFlag = TRUE;
	m_readWaveDataSize = 0;

	//ÉfÉRÅ[ÉhèÄîı
	//decodePos		= 0;		//éûä‘éÊìæóp

	::SetFilePointer(m_hFile, m_firstSync, NULL, FILE_BEGIN);
	if ( m_waveFlag == FALSE ){
		m_inputSize		= m_decInfo.maxInputSize * m_inputBlock;
		m_inputBuf		= (byte*)::realloc(m_inputBuf, m_inputSize);
		m_inputPtr		= m_inputBuf;
		m_inputBytes	= input_read(m_inputPtr, m_inputSize);
		if (m_inputBytes == 0) {
			m_errNo = ERR_INVALID_VALUE;
			return FALSE;
		}
	}

	//èoóÕÉfÉoÉCÉXÇäJÇ≠
	OutputInfo	outputInfo;
	outputInfo.channels			= m_decInfo.channels;
	outputInfo.bitsPerSample	= m_decInfo.bitsPerSample;
	outputInfo.frequency		= m_decInfo.frequency;

	if ( m_waveFlag == FALSE ){
		outputInfo.bufferSize		= m_decInfo.outputSize * m_outputBlock_in;
	} else {
		m_outWaveBufSize = m_decInfo.outputSize * m_outputBlock_in_wav;
		m_inWaveBufSize = m_decInfo.outputSize;
		outputInfo.bufferSize		= m_outWaveBufSize;
	}

	if (!output_open(NULL, outputInfo)) {
		m_errNo = ERR_OPEN_OUT_DEVICE;
		return FALSE;
	}

	//m_outputBuf_in= (BYTE*)::realloc(m_outputBuf_in, m_decInfo.outputSize * m_outputBlock_in);
	m_outputBuf_in= (BYTE*)::realloc(m_outputBuf_in, outputInfo.bufferSize);

	if ( m_waveFlag == FALSE ){
		if (!::mp3DecodeStart(m_inputPtr, m_inputBytes)) {
			output_close();
			m_errNo = ERR_DECODE;
			return FALSE;
		}
		m_inputPtr += m_decInfo.skipSize;
		m_inputBytes -= m_decInfo.skipSize;
	}

	//ÉfÉRÅ[ÉhópÉXÉåÉbÉhäJén
	m_hThread = (HANDLE)::_beginthreadex(NULL, 0, 
			(THREADFUNC)input_ThreadProc, NULL, 0, (unsigned int*)&m_threadID);
	if (m_hThread == NULL) {
		output_close();
		m_errNo = ERR_DECODE_THREAD;
		return FALSE;
	}

	//èÛë‘É`ÉFÉbÉN
	int Cnt;
	for ( Cnt = 0; Cnt < 100; Cnt++){
		::Sleep(100);
		if ( m_state == STATE_PLAY ) {
			break;
		} else if ( Cnt == 99 ){
			m_errNo = ERR_PLAY;
			return FALSE;
		}
	}

	return TRUE;
}

BOOL input_decodeWave(const char* pszWaveName)
{
	BYTE			buf[1024];
	BOOL			result;
	DWORD			bytesWrite;

	if (m_state != STATE_STOP) {
		//í‚é~íÜà»äOÇÕé¿çsïsâ¬
		m_errNo = ERR_NOT_STATE_STOP;
		return FALSE;
	}

	decFlag = 1;

	//ÉtÉ@ÉCÉãÉIÅ[ÉvÉì
	m_hWaveFile = ::CreateFile(pszWaveName, GENERIC_READ + GENERIC_WRITE, FILE_SHARE_READ,
			NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if (m_hWaveFile == INVALID_HANDLE_VALUE) {
		m_errNo = ERR_WAV_FILE_OPEN;
		return FALSE;
	}

	SetFilePointer(m_hWaveFile, 0, NULL, FILE_BEGIN);
	::strcpy((char*)buf,"RIFF");
	result = ::WriteFile(m_hWaveFile, &buf, 4, &bytesWrite, NULL);
	::memset(&buf, 0, 4);
	result = ::WriteFile(m_hWaveFile, &buf, 4, &bytesWrite, NULL);
	::strcpy((char*)buf,"WAVEfmt ");
	result = ::WriteFile(m_hWaveFile, &buf, 8, &bytesWrite, NULL);

	//É`ÉÉÉìÉNÉTÉCÉY
	buf[0] = 0x10;
	buf[1] = 0x00;
	buf[2] = 0x00;
	buf[3] = 0x00;
	result = ::WriteFile(m_hWaveFile, &buf, 4, &bytesWrite, NULL);

	//ÉtÉHÅ[É}ÉbÉgÇhÇc
	buf[0] = 0x01;
	buf[1] = 0x00;
	result = ::WriteFile(m_hWaveFile, &buf, 2, &bytesWrite, NULL);

	//É`ÉÉÉìÉlÉãêî
	buf[0] = m_decInfo.channels;
	buf[1] = 0x00;
	result = ::WriteFile(m_hWaveFile, &buf, 2, &bytesWrite, NULL);

	//ÉTÉìÉvÉäÉìÉOÉåÅ[Ég
	DWORD nSamplesPerSec = m_decInfo.frequency;
	result = ::WriteFile(m_hWaveFile, &nSamplesPerSec, 4, &bytesWrite, NULL);

	//ì]ëóÉåÅ[Ég
	WORD  nBlockAlign = (m_decInfo.channels * m_decInfo.bitsPerSample / 8);
	DWORD nAvgBytesPerSec = (nBlockAlign * m_decInfo.frequency);
	result = ::WriteFile(m_hWaveFile, &nAvgBytesPerSec, 4, &bytesWrite, NULL);

	//ÉuÉçÉbÉNÉTÉCÉYêî
	result = ::WriteFile(m_hWaveFile, &nBlockAlign, 2, &bytesWrite, NULL);

	//ó éqâªÉrÉbÉgêî
	buf[0] = m_decInfo.bitsPerSample;
	buf[1] = 0x00;
	result = ::WriteFile(m_hWaveFile, &buf, 2, &bytesWrite, NULL);

	::strcpy((char*)buf,"data");
	result = ::WriteFile(m_hWaveFile, &buf, 4, &bytesWrite, NULL);
	::memset(&buf, 0, 4);
	result = ::WriteFile(m_hWaveFile, &buf, 4, &bytesWrite, NULL);

	//ÉfÉRÅ[ÉhèÄîı
	//decodePos		= 0;		//éûä‘éÊìæóp
	::SetFilePointer(m_hFile, m_firstSync, NULL, FILE_BEGIN);
	m_inputSize		= m_decInfo.maxInputSize * m_inputBlock;
	m_inputBuf		= (byte*)::realloc(m_inputBuf, m_inputSize);
	m_inputPtr		= m_inputBuf;
	m_inputBytes	= input_read(m_inputPtr, m_inputSize);

	if (m_inputBytes == 0) {
		m_errNo = ERR_INVALID_VALUE;
		return FALSE;
	}

	//èoóÕÉfÉoÉCÉXÇäJÇ≠
/*
	OutputInfo	outputInfo;
	outputInfo.channels			= m_decInfo.channels;
	outputInfo.bitsPerSample	= m_decInfo.bitsPerSample;
	outputInfo.frequency		= m_decInfo.frequency;
	outputInfo.bufferSize		= m_decInfo.maxOutputSize * m_outputBlock;

	if (!output_open(NULL, outputInfo)) {
		return FALSE;
	}
*/
	m_outputBuf_in= (BYTE*)::realloc(m_outputBuf_in, m_decInfo.outputSize);

	if (!::mp3DecodeStart(m_inputPtr, m_inputBytes)) {
		output_close();
		m_errNo = ERR_DECODE;
		return FALSE;
	}

	m_inputPtr += m_decInfo.skipSize;
	m_inputBytes -= m_decInfo.skipSize;

	//ÉfÉRÅ[ÉhópÉXÉåÉbÉhäJén
	m_hThread = (HANDLE)::_beginthreadex(NULL, 0, 
			(THREADFUNC)input_ThreadProc, NULL, 0, (unsigned int*)&m_threadID);
	if (m_hThread == NULL) {
		output_close();
		m_errNo = ERR_DECODE_THREAD;
		return FALSE;
	}

	//èÛë‘É`ÉFÉbÉN
	int Cnt;
	for ( Cnt = 0; Cnt < 100; Cnt++){
		::Sleep(100);
		if ( m_state == STATE_PLAY ) {
			break;
		} else if ( Cnt == 99 ){
			m_errNo = ERR_PLAY;
			return FALSE;
		}
	}

	return TRUE;
}

BOOL input_playDecodeWave(const char* pszWaveName)
{
	BYTE			buf[1024];
	BOOL			result;
	DWORD			bytesWrite;

	if (m_state != STATE_STOP) {
		//í‚é~íÜà»äOÇÕé¿çsïsâ¬
		m_errNo = ERR_NOT_STATE_STOP;
		return FALSE;
	}

	//WAVEÉtÉ@ÉCÉã
	if ( m_waveFlag == TRUE ){
		//m_errNo = ERR_NOT_STATE_STOP;
		return FALSE;
	}

	decFlag = 2;
	m_fadeOutFlag = false;
	m_playFlag = TRUE;
	m_readWaveDataSize = 0;

	::SetFilePointer(m_hFile, m_firstSync, NULL, FILE_BEGIN);
	m_inputSize		= m_decInfo.maxInputSize * m_inputBlock;
	m_inputBuf		= (byte*)::realloc(m_inputBuf, m_inputSize);
	m_inputPtr		= m_inputBuf;
	m_inputBytes	= input_read(m_inputPtr, m_inputSize);
	if (m_inputBytes == 0) {
		m_errNo = ERR_INVALID_VALUE;
		return FALSE;
	}

	//èoóÕÉfÉoÉCÉXÇäJÇ≠
	OutputInfo	outputInfo;
	outputInfo.channels			= m_decInfo.channels;
	outputInfo.bitsPerSample	= m_decInfo.bitsPerSample;
	outputInfo.frequency		= m_decInfo.frequency;

	outputInfo.bufferSize		= m_decInfo.outputSize * m_outputBlock_in;

	if (!output_open(NULL, outputInfo)) {
		m_errNo = ERR_OPEN_OUT_DEVICE;
		return FALSE;
	}

	m_outputBuf_in= (BYTE*)::realloc(m_outputBuf_in, outputInfo.bufferSize);

	if (!::mp3DecodeStart(m_inputPtr, m_inputBytes)) {
		output_close();
		m_errNo = ERR_DECODE;
		return FALSE;
	}
	m_inputPtr += m_decInfo.skipSize;
	m_inputBytes -= m_decInfo.skipSize;

	//ÉtÉ@ÉCÉãÉIÅ[ÉvÉì
	m_hWaveFile = ::CreateFile(pszWaveName, GENERIC_READ + GENERIC_WRITE, FILE_SHARE_READ,
			NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if (m_hWaveFile == INVALID_HANDLE_VALUE) {
		m_errNo = ERR_WAV_FILE_OPEN;
		return FALSE;
	}

	SetFilePointer(m_hWaveFile, 0, NULL, FILE_BEGIN);
	::strcpy((char*)buf,"RIFF");
	result = ::WriteFile(m_hWaveFile, &buf, 4, &bytesWrite, NULL);
	::memset(&buf, 0, 4);
	result = ::WriteFile(m_hWaveFile, &buf, 4, &bytesWrite, NULL);
	::strcpy((char*)buf,"WAVEfmt ");
	result = ::WriteFile(m_hWaveFile, &buf, 8, &bytesWrite, NULL);

	//É`ÉÉÉìÉNÉTÉCÉY
	buf[0] = 0x10;
	buf[1] = 0x00;
	buf[2] = 0x00;
	buf[3] = 0x00;
	result = ::WriteFile(m_hWaveFile, &buf, 4, &bytesWrite, NULL);

	//ÉtÉHÅ[É}ÉbÉgÇhÇc
	buf[0] = 0x01;
	buf[1] = 0x00;
	result = ::WriteFile(m_hWaveFile, &buf, 2, &bytesWrite, NULL);

	//É`ÉÉÉìÉlÉãêî
	buf[0] = m_decInfo.channels;
	buf[1] = 0x00;
	result = ::WriteFile(m_hWaveFile, &buf, 2, &bytesWrite, NULL);

	//ÉTÉìÉvÉäÉìÉOÉåÅ[Ég
	DWORD nSamplesPerSec = m_decInfo.frequency;
	result = ::WriteFile(m_hWaveFile, &nSamplesPerSec, 4, &bytesWrite, NULL);

	//ì]ëóÉåÅ[Ég
	WORD  nBlockAlign = (m_decInfo.channels * m_decInfo.bitsPerSample / 8);
	DWORD nAvgBytesPerSec = (nBlockAlign * m_decInfo.frequency);
	result = ::WriteFile(m_hWaveFile, &nAvgBytesPerSec, 4, &bytesWrite, NULL);

	//ÉuÉçÉbÉNÉTÉCÉYêî
	result = ::WriteFile(m_hWaveFile, &nBlockAlign, 2, &bytesWrite, NULL);

	//ó éqâªÉrÉbÉgêî
	buf[0] = m_decInfo.bitsPerSample;
	buf[1] = 0x00;
	result = ::WriteFile(m_hWaveFile, &buf, 2, &bytesWrite, NULL);

	::strcpy((char*)buf,"data");
	result = ::WriteFile(m_hWaveFile, &buf, 4, &bytesWrite, NULL);
	::memset(&buf, 0, 4);
	result = ::WriteFile(m_hWaveFile, &buf, 4, &bytesWrite, NULL);

	//ÉfÉRÅ[ÉhópÉXÉåÉbÉhäJén
	m_hThread = (HANDLE)::_beginthreadex(NULL, 0, 
			(THREADFUNC)input_ThreadProc, NULL, 0, (unsigned int*)&m_threadID);
	if (m_hThread == NULL) {
		output_close();
		m_errNo = ERR_DECODE_THREAD;
		return FALSE;
	}

	//èÛë‘É`ÉFÉbÉN
	int Cnt;
	for ( Cnt = 0; Cnt < 100; Cnt++){
		::Sleep(100);
		if ( m_state == STATE_PLAY ) {
			break;
		} else if ( Cnt == 99 ){
			m_errNo = ERR_PLAY;
			return FALSE;
		}
	}

	return TRUE;
}

BOOL input_stop()
{
	if (m_state == STATE_STOP) {
		//í‚é~íÜÇÕèàóùÇµÇ»Ç¢
		m_errNo = ERR_STATE_STOP;
		return TRUE;
	}

	if ( m_threadID == 0 ) {
		output_stop();
	//	output_close();
	}
	if (::PostThreadMessage(m_threadID, EV_STOP, 0, 0)) {
		::WaitForSingleObject(m_hThread, 500);
	}

	m_playSec = 0;
	m_playSecCur = 0;
	m_playMs = 0;
	m_firstSample = 0;

	//èÛë‘É`ÉFÉbÉN
	int Cnt;
	for ( Cnt = 0; Cnt < 100; Cnt++){
		if ( m_state == STATE_STOP ) {
			break;
		} else if ( Cnt == 99 ){
			m_errNo = ERR_STOP;
			return FALSE;
		}
		::Sleep(10);
	}

	return TRUE;
}

BOOL input_pause()
{
	int w_state = m_state;
	if (w_state == STATE_PAUSE ||
			w_state == STATE_STOP) {
		//àÍéûí‚é~íÜÅAí‚é~íÜÇÕèàóùÇµÇ»Ç¢
		m_errNo = ERR_NOT_STATE_PLAY;
		return TRUE;
	}

	if ( m_threadID == 0 ) {
		input_actionPause();
		return TRUE;
	} else {
		return input_sendEvent(EV_PAUSE, 0);
	}
}

BOOL input_restart()
{
	if (m_state != STATE_PAUSE) {
		//àÍéûí‚é~íÜà»äOÇÕèàóùÇµÇ»Ç¢
		return TRUE;
	}

	if ( m_threadID == 0 ) {
		input_actionRestart();
		return TRUE;
	} else {
		return input_sendEvent(EV_RESTART, 0);
	}
}

void input_setEqualizer(int* pTable)
{
	if (m_threadID) {
		input_sendEvent(EV_EQUALIZER, (int)pTable);
	}
	else {
		mp3SetEqualizer(pTable);
	}
}


void input_seekByFrame(int frame)
{
	m_seekFlag = TRUE;

	int sample = input_bytesToSample(frame * m_decInfo.outputSize);
	input_seekBySample(sample);
}

void input_seekBySample(int sample)
{
	int w_state = m_state;
	m_seekFlag = TRUE;
	if ( m_threadID == 0 && m_state != STATE_STOP) {
		input_stop();
		//::Sleep(100);
		input_play();
		if ( w_state == STATE_PAUSE )
			input_pause();
		//::Sleep(100);
		//input_actionSeek(sample);
	}
	//} else {
	//if ( m_threadID != 0 )
		input_sendEvent(EV_SEEK, sample);
	//}
	output_softVolumeMax();
}

void input_seekBySec(int sec)
{
	m_seekFlag = TRUE;
	int sample = input_secToSample((int)(sec / m_playSpeed));
	input_seekBySample(sample);
}

int input_FrameSizeByte(void)
{
	return 144 * m_decInfo.bitRate / m_decInfo.frequency  + m_decInfo.header.padding;
}

int input_bytesToSample(int bytes)
{
	long OutPutSample = 0;

 
	if (m_decInfo.bitsPerSample != 0 && m_decInfo.channels != 0){
		if ( bytes > ALLOW_DATA_SIZE ) {
			OutPutSample = (bytes / m_decInfo.bitsPerSample / m_decInfo.channels * 8);
		} else {
			OutPutSample = (bytes * 8 / m_decInfo.bitsPerSample / m_decInfo.channels);
		}
	} else {
		OutPutSample = 0;
	}

	return OutPutSample;
	
}

int input_secToSample(int sec)
{
	return (sec * m_decInfo.frequency);
}

int input_sampleToSec(int sample)
{
	if ( m_decInfo.frequency == 0 ) return 0;
	return (sample / m_decInfo.frequency);
}

long input_sampleToMs(int sample)
{
	if ( m_decInfo.frequency == 0 ) return 0;
	return (sample * 1000 / m_decInfo.frequency);
}

BOOL input_skipNextSync()
{
	int sync;
	m_inputPtr   ++;
	m_inputBytes --;
	for (;;) {
		if (::mp3FindSync(m_inputPtr, m_inputBytes, &sync)) {
			m_inputPtr += sync;
			m_inputBytes -= sync;
			return true;
		}
		::memmove(m_inputBuf, m_inputPtr, m_inputBytes);
		int bytesRead = input_read(&m_inputBuf[m_inputBytes], m_inputSize - m_inputBytes);
		if (bytesRead == 0) {
			m_errNo = ERR_MP3_FILE_READ;
			return false;//no more data
		}
		m_inputPtr = m_inputBuf;
		m_inputBytes += bytesRead;
	}
}

/*
BOOL input_seek(int pos)
{
	int w_state = m_state;
	if (w_state != STATE_PLAY &&
			w_state != STATE_PAUSE) {
		//çƒê∂íÜÅAàÍéûí‚é~íÜÇÃÇ›èàóù
		return FALSE;
	}

	if ( m_threadID == 0 ) {
		int sample = input_secToSample(pos);
		input_actionSeek(sample);
		return TRUE;
	} else {
		return input_sendEvent(EV_SEEK, pos);
	}

	return TRUE;
}
*/

BOOL input_setVolume(int left, int right)
{
	return output_setVolume(left, right);
}

BOOL input_getVolume(int* left, int* right)
{
	return output_getVolume(left, right);
}

int input_getWaveOutSupport()
{
    int support = output_getWaveOutSupport();
    if (support & WAVECAPS_LRVOLUME) {
        //ç∂âEÉnÅ[ÉhÉ{ÉäÉÖÅ[ÉÄê›íËâ¬î
		return 2;
    }
    else if (support & WAVECAPS_VOLUME) {
        //ÉnÅ[ÉhÉ{ÉäÉÖÅ[ÉÄê›íËâ¬î\Åiç∂âEí≤êÆïsâ¬Åj
		return 1;
    }
    else {
        //ÉnÅ[ÉhÉ{ÉäÉÖÅ[ÉÄê›íËïsâ¬î
		return 0;
    }
}


void input_setSoftVolume(int left, int right)
{
	output_setSoftVolume(left, right);
}

void input_getSoftVolume(int* left, int* right)
{
	output_getSoftVolume(left, right);
}

void input_getSpectrum(int* pSpecL, int* pSpecR)
{
	output_getSpectrum(output_getCurrentSample(), pSpecL, pSpecR);
	return;
}

void input_getWave(int* pWaveL, int* pWaveR)
{
	DWORD sample = output_getCurrentSample();
	if ( sample ) output_getWave(sample, pWaveL, pWaveR);
	return;
}

BOOL input_sendEvent(int code, int param)
{
	HANDLE hEvent = ::CreateEvent(NULL, false, false, NULL);
	if (hEvent == NULL) {
		m_errNo = ERR_CREATE_EVENT;
		return FALSE;
	}
	if (::PostThreadMessage(m_threadID, code, (WPARAM)hEvent, param)) {
		::WaitForSingleObject(hEvent, 500);
	}
	::CloseHandle(hEvent);
	return TRUE;
}

int input_ThreadProc()
{
	return input_decode();
}

int input_decode()
{
	MPEG_DECODE_PARAM param;
	MSG		msg;
	DWORD	readBytes;
	int		oldSec;
	int		waitLCount = 0;
	int		waitRCount = 0;

	BOOL			result;
	DWORD			bytesWrite;


	oldSec = 0;
	m_state = STATE_PLAY;
	m_playSec = 0;
	m_playSecCur = 0;
	m_playMs = 0;
	m_playSamples = 0;
	m_firstSample = 0;
	decodePos = 0;
	param.outputBuf	= m_outputBuf_in;

	if ( callBackFlag == TRUE )
		m_pProc(INPUT_MSG_PLAY, m_playSec);
	
	for (;;) {
		for (;;) {
			if ( m_waveFlag == FALSE ){
				param.inputBuf	= m_inputPtr;
				param.inputSize	= m_inputBytes;

				if (!::mp3DecodeFrame(&param)) {
					int err = ::mp3GetLastError();
					if (err == MP3_ERROR_OUT_OF_BUFFER) {
						break;//read next buffer
					}
					//invalid frame !!
					if ( callBackFlag == TRUE )
							m_pProc(INPUT_MSG_ERROR, 0);
					if (!input_skipNextSync()) {
						goto exit;
					}
					continue;
				}

				m_bitRate = param.bitRate;
				//decodePos    += param.inputSize;		//éûä‘éÊìæóp
				m_inputPtr   += param.inputSize;
				m_inputBytes -= param.inputSize;
				//m_playMs     += m_msPerFrame;
			} else {
				//WAVEÉtÉ@ÉCÉã
				if ( (int)dataSize >= m_readWaveDataSize ) {
					param.outputSize = m_inWaveBufSize;
					m_readWaveDataSize = m_readWaveDataSize + param.outputSize;
					if ( (int)dataSize < m_readWaveDataSize ) param.outputSize = m_readWaveDataSize - dataSize;
					if ( m_bitRate == 8 ) {
						::memset(param.outputBuf, 128, param.outputSize);
					} else {
						::memset(param.outputBuf, 0, param.outputSize);
					}
					readBytes = input_read(param.outputBuf, param.outputSize);
					if (readBytes == 0) {
						goto exit;//no more data
					}
				} else {
					goto exit;
				}
			}

			//ÉfÉRÅ[Éh or WAVEÉtÉ@ÉCÉãèoóÕ
			if (decFlag == 0 || decFlag == 2) {
				if (output_write(param.outputBuf, param.outputSize)){
					m_playSamples = output_getCurrentSample();
					//m_playMs = (((float)m_playSamples + (float)m_firstSample) * (float)1000 / (float)m_decInfo.frequency) / (float)1000;
					m_playSec = (int)(input_sampleToSec(m_playSamples + m_firstSample) * m_playSpeed);
					if (m_playSec != oldSec) {
						if ( callBackFlag == TRUE )
							m_pProc(INPUT_MSG_PLAY, m_playSec);
						oldSec = m_playSec;
					}
				}
			}
			if (decFlag == 1 || decFlag == 2) {
				decodePos += param.outputSize;
				result = ::WriteFile(m_hWaveFile, param.outputBuf, param.outputSize, &bytesWrite, NULL);
				m_playSamples = output_getCurrentSample();
				//m_playMs = (((float)m_playSamples + (float)m_firstSample) * (float)1000 / (float)m_decInfo.frequency) / (float)1000;
				m_playSec = (int)(input_sampleToSec(m_playSamples + m_firstSample) * m_playSpeed);
				if (m_playSec != oldSec) {
					if ( callBackFlag == TRUE )
						m_pProc(INPUT_MSG_PLAY, m_playSec);
					oldSec = m_playSec;
				}
			}

			if ( m_fadeOutFlag ){
				waitLCount++;
				waitRCount++;

				if ( (m_lvol > 50 && waitLCount == 2) ||
					 (m_lvol > 10 && waitLCount == 3) ||
					 (m_lvol > 5  && waitLCount == 5) ||
					 (m_lvol > 0  && waitLCount == 10) ) {
					m_lvol--;
					waitLCount=0;
				}

				if ( (m_rvol > 50 && waitRCount == 2) ||
					 (m_rvol > 10 && waitRCount == 3) ||
					 (m_rvol > 5  && waitRCount == 5) ||
					 (m_rvol > 0  && waitRCount == 10) ) {
					m_rvol--;
					waitRCount=0;
				}
				output_setVolume(m_lvol, m_rvol);

				if (m_lvol == 0 && m_rvol == 0) {
					m_lvol=m_olvol;
					m_rvol=m_orvol;
					m_fadeOutFlag = false;
					output_stop();
					goto exit;
				}
			}
			
			if (::PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
				input_handleEvent(msg);
			}
			while (m_state == STATE_PAUSE) {
				if (::GetMessage(&msg, NULL, 0, 0)) {
					input_handleEvent(msg);
				}
			}
			if (m_state == STATE_STOP) {
				goto exit;
			}
		}
		if ( m_waveFlag == FALSE ){
			::memmove(m_inputBuf, m_inputPtr, m_inputBytes);
			readBytes = input_read(&m_inputBuf[m_inputBytes], m_inputSize - m_inputBytes);
			if (readBytes == 0) break;//no more data
			m_inputPtr = m_inputBuf;
			m_inputBytes += readBytes;
			::Sleep(m_inputSleep);
		}
	}
	
exit:

	m_threadID = 0;

	while (::PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
		HANDLE hEvent = (HANDLE)msg.wParam;
		if (hEvent != NULL) {
			::SetEvent(hEvent);
		}
	}
	::CloseHandle(m_hThread);
	m_hThread = NULL;
	if (m_state != STATE_STOP) {
		input_actionStop(TRUE);
		if ( callBackFlag == TRUE )
			m_pProc(INPUT_MSG_PLAYDONE, 0);
	}

	m_playSec = 0;
	m_playSecCur = 0;
	m_playMs = 0;
	m_playSamples = 0;
	m_firstSample = 0;

	if (decFlag == 1 || decFlag == 2) {
		::SetEndOfFile(m_hWaveFile);
		DWORD cSize = ::GetFileSize(m_hWaveFile, NULL) - 8;
		SetFilePointer(m_hWaveFile, 4, NULL, FILE_BEGIN);
		result = ::WriteFile(m_hWaveFile, &cSize, 4, &bytesWrite, NULL);

		DWORD dSize = cSize - 36;
		SetFilePointer(m_hWaveFile, 40, NULL, FILE_BEGIN);
		result = ::WriteFile(m_hWaveFile, &dSize, 4, &bytesWrite, NULL);

		::CloseHandle(m_hWaveFile);
	}
	return 0;
}

void input_handleEvent(const MSG& msg)
{
	switch (msg.message) {
	case EV_STOP:
		input_actionStop(FALSE);
		break;
	case EV_PAUSE:
		input_actionPause();
		m_state = STATE_PAUSE;
		break;
	case EV_RESTART:
		input_actionRestart();
		m_state = STATE_PLAY;
		break;
	case EV_SEEK:
		input_actionSeek(msg.lParam);
		break;
	case EV_EQUALIZER:
		mp3SetEqualizer((int*)msg.lParam);
		break;
	}
	HANDLE hEvent = (HANDLE)msg.wParam;
	if (hEvent != NULL) {
		::SetEvent(hEvent);
	}

}

void input_actionStop(BOOL wait)
{
	if (!wait) {
		//ã≠êßí‚é~
		output_stop();
	}
	output_close();
	if (wait) {
		//âï˙
		output_stop();
	}
	m_state = STATE_STOP;
	if ( callBackFlag == TRUE && !wait)
		m_pProc(INPUT_MSG_STOP, 0);
}

void input_actionPause()
{
	if (m_state == STATE_STOP) {
		return;//ignore
	}
	if (m_state == STATE_PLAY) {
		output_pause();
	}
	m_state = STATE_PAUSE;
	if ( callBackFlag == TRUE )
		m_pProc(INPUT_MSG_PAUSE, 0);
}

void input_actionRestart()
{
	if (m_state == STATE_PAUSE) {
		output_restart();
	}
	m_state = STATE_PLAY;
	if ( callBackFlag == TRUE )
		m_pProc(INPUT_MSG_PLAY, m_playSec);

}

void input_actionSeek(int seekSample)
{

//	input_actionSpeedSeek(seekSample);
//	return;
	MPEG_DECODE_PARAM param;
	int bytesRead, outputBytes;

	output_stop();
	m_state = STATE_SEEK;
	::SetFilePointer(m_hFile, m_firstSync, NULL, FILE_BEGIN);
	
	m_playSecCur = 0;
	m_playMs = 0;
	m_inputBytes = 0;
	m_playSamples = 0;
	outputBytes = 0;
	param.outputBuf	= m_outputBuf_in;
	m_readWaveDataSize = 0; 

	if ( m_waveFlag == FALSE){
		::mp3MuteStart(&param);
		for (;;) {
			::memmove(m_inputBuf, m_inputPtr, m_inputBytes);
			bytesRead = input_read(&m_inputBuf[m_inputBytes], m_inputSize - m_inputBytes);
			if (bytesRead == 0) break;//no more data
			m_inputPtr = m_inputBuf;
			m_inputBytes += bytesRead;

			for (;;) {
				param.inputBuf	= m_inputPtr;
				param.inputSize	= m_inputBytes;
				if (!::mp3DecodeFrame(&param)) {
					int err = ::mp3GetLastError();
					if (err == MP3_ERROR_OUT_OF_BUFFER) {
						break;//read next buffer
					}
					//invalid frame !!
					if ( callBackFlag == TRUE )
						m_pProc(INPUT_MSG_ERROR, 0);
					if (!input_skipNextSync()) {
						goto exit;
					}
					continue;
				}

				m_inputPtr   += param.inputSize;
				m_inputBytes -= param.inputSize;
				outputBytes += param.outputSize;
				m_firstSample = input_bytesToSample(outputBytes);
				if (m_firstSample >= seekSample) {
					m_state = STATE_PLAY;
					goto exit;
				}
			}
		}
	} else {
		m_firstSample = seekSample;
		decodePos = m_firstSample;
		outputBytes = (int)output_sampleToBytes(m_firstSample);
		::SetFilePointer(m_hFile, outputBytes, NULL, FILE_CURRENT);
		m_readWaveDataSize = m_readWaveDataSize + outputBytes;
		m_state = STATE_PLAY;
	}

exit:
	if ( m_waveFlag == FALSE ) 
		::mp3MuteEnd(&param);
	m_seekFlag = FALSE;

//	if ( callBackFlag == TRUE )
//		m_pProc(INPUT_MSG_PLAY, m_playSec);

}

long input_sampleToMp3(int Sample) {

	int Sec;
	int xBit;
	Sec = input_sampleToSec(Sample);
	xBit = Sec * m_bitRate;
	return xBit / 8;

}

__int64 myFileSeek (HANDLE hf, __int64 distance, DWORD MoveMethod)
{
   LARGE_INTEGER li;

   li.QuadPart = distance;

   li.LowPart = SetFilePointer (hf, li.LowPart, &li.HighPart, MoveMethod);

/*/if (li.LowPart == INVALID_SET_FILE_POINTER && GetLastError() != NO_ERROR)
   {
      li.QuadPart = -1;
   }/*/

   return li.QuadPart;
}

/*/
void input_actionSpeedSeek(int seekSample)
{
	MPEG_DECODE_PARAM param;
	DWORD bytesRead, outputBytes;
	long NextBufRead;

	output_stop();
	m_state = STATE_SEEK;
	NextBufRead = input_sampleToMp3(seekSample);
	//NextBufRead = 71680000;
	//::SetFilePointer(m_hFile, m_firstSync + NextBufRead, NULL, FILE_BEGIN);
		

//X-MaD 02/08/01
	BYTE buf[2048];
	int sync;
	long result;
	MPEG_DECODE_INFO decInfo;
	bool PORCODIO = false;
	int InputBytes = sizeof(buf);
	while (!PORCODIO) {

	
	::SetFilePointer(m_hFile, NextBufRead, NULL, FILE_BEGIN);
	result = ::ReadFile(m_hFile, buf, InputBytes, &bytesRead, NULL);
	if (!result || bytesRead == 0) {
		m_errNo = ERR_MP3_FILE_READ;
		::input_close();
		return;
	}
	
	
	if (!::mp3FindSync(buf, InputBytes, &sync)) {
		::input_close();
		m_errNo = ERR_FRAME_HEADER_NOT_FOUND;
		return;
	}

	NextBufRead += sync;

	::SetFilePointer(m_hFile, NextBufRead, NULL, FILE_BEGIN);
	result = ::ReadFile(m_hFile, buf, sizeof(buf), &bytesRead, NULL);
	if (!result || bytesRead == 0) {
		::input_close();
		m_errNo = ERR_MP3_FILE_READ;
		return;
	}


	if (!::mp3GetDecodeInfo(buf, bytesRead, &decInfo, 1)) {
		m_errNo = ERR_FRAME_HEADER_READ;
		NextBufRead += InputBytes;
	}
	else {
		PORCODIO = true;
	}

	}

::SetFilePointer(m_hFile, NextBufRead, NULL, FILE_BEGIN);
//X-MaD 02/08/01

	m_playSecCur = 0;
	m_playMs = 0;
	m_inputBytes = 0;
	m_playSamples = 0;
	outputBytes = 0;
	param.outputBuf	= m_outputBuf_in;
	m_readWaveDataSize = 0; 

	if ( m_waveFlag == FALSE){
		::mp3MuteStart(&param);
		for (;;) {
			::memmove(m_inputBuf, m_inputPtr, m_inputBytes);
			bytesRead = input_read(&m_inputBuf[m_inputBytes], m_inputSize - m_inputBytes);
			if (bytesRead == 0) break;//no more data
			m_inputPtr = m_inputBuf;
			m_inputBytes += bytesRead;

			for (;;) {
				param.inputBuf	= m_inputPtr;
				param.inputSize	= m_inputBytes;
				if (!::mp3DecodeFrame(&param)) {
					int err = ::mp3GetLastError();
					if (err == MP3_ERROR_OUT_OF_BUFFER) {
						break;//read next buffer
					}
					//invalid frame !!
					if ( callBackFlag == TRUE )
						m_pProc(INPUT_MSG_ERROR, 0);
					if (!input_skipNextSync()) {
						goto exit;
					}
					continue;
				}

				m_inputPtr   += param.inputSize;
				m_inputBytes -= param.inputSize;
				outputBytes += param.outputSize;
				m_firstSample = input_bytesToSample(outputBytes);
				if (m_firstSample >= seekSample) {
					m_state = STATE_PLAY;
					goto exit;
				}
			}
		}
	} else {
		m_firstSample = seekSample;
		decodePos = m_firstSample;
		outputBytes = (int)output_sampleToBytes(m_firstSample);
		::SetFilePointer(m_hFile, outputBytes, NULL, FILE_CURRENT);
		m_readWaveDataSize = m_readWaveDataSize + outputBytes;
	}

exit:
	if ( m_waveFlag == FALSE ) 
		::mp3MuteEnd(&param);
	m_seekFlag = FALSE;

//	if ( callBackFlag == TRUE )
//		m_pProc(INPUT_MSG_PLAY, m_playSec);

}/*/

/******************** RIFF.cpp ************************/

BOOL readRIFF(HANDLE FileHandle, int* fSync, DWORD* dSize)
{
	RiffChunk	riff;
	CommonChunk	common;
	BOOL		result;
	DWORD		bytesRead;
	DWORD       wkDataSize;
	DWORD       wkDataPoint;

	RiffStatus = 0;

	//ÉtÉ@ÉCÉãÇÃì™Ç…ñﬂÇËÅARIFFÉ`ÉÉÉìÉNÇì«Ç›çûÇﬁ
	::SetFilePointer(FileHandle, 0, NULL, FILE_BEGIN);
	result = ::ReadFile(FileHandle, &riff, sizeof(riff), &bytesRead, NULL);
	if (!result || bytesRead != sizeof(riff)) {
		return FALSE;
	}
	if (::memcmp(riff.chunkID, "RIFF", 4) != 0) {
		return FALSE;
	}

	//FileType
	if (memcmp(riff.formType, "WAVE", 4) == 0) {
		m_FileType += 1;
	}
	if (memcmp(riff.formType, "RMP3", 4) == 0) {
		m_FileType += 2;
	}

	*fSync = sizeof(riff);

	//dataÉ`ÉÉÉìÉNÇíTÇ∑
	for (;;) {
		result = ::ReadFile(FileHandle, &common, sizeof(common), &bytesRead, NULL);
		if (!result || bytesRead != sizeof(common)) {
			return FALSE;
		}

		if (::memcmp(common.chunkID, "fmt ", 4) == 0) {
			//Mpegà»äOÇÃÉtÉHÅ[É}ÉbÉgÉ`ÉFÉbÉN
			//::ReadFile(FileHandle, &m_waveFormat, common.chunkSize, &bytesRead, NULL);
			::ReadFile(FileHandle, &m_waveFormat, sizeof(WAVEFORMATEX), &bytesRead, NULL);
			if (!result || bytesRead != sizeof(WAVEFORMATEX)) {
				RiffStatus = -1;
				return FALSE;
			}
			//ÉtÉHÅ[É}ÉbÉgTagÉ`ÉFÉbÉN
			if (m_waveFormat.wFormatTag == 85 || m_waveFormat.wFormatTag == 80) {
				//MPEG Audio ÉfÅ[É^
				//common.chunkSize = 0;  //É`ÉÉÉìÉNÇì«Ç›îÚÇŒÇ≥Ç»Ç¢
			} else if (m_waveFormat.wFormatTag == 1) {
				//ÉäÉAÉãÇoÇbÇlÉfÅ[É^
				RiffStatus = 1;
				//common.chunkSize = 0;  //É`ÉÉÉìÉNÇì«Ç›îÚÇŒÇ≥Ç»Ç¢
			} else {
				//ÇªÇÃëºÇÃå`éÆ
				RiffStatus = -1;
				return FALSE; 
			}
			::SetFilePointer(FileHandle, -(long)sizeof(WAVEFORMATEX), NULL, FILE_CURRENT);
/*
			if (common.chunkSize != 0x1E && common.chunkSize != 0x20) {
				if (common.chunkSize == 0x10 || common.chunkSize == 0x12) {
					::ReadFile(FileHandle, &m_waveFormat, common.chunkSize, &bytesRead, NULL);
					if (!result || bytesRead != common.chunkSize) {
						RiffStatus = -1;
						return FALSE;
					}
					RiffStatus = 1;
					common.chunkSize = 0;  //É`ÉÉÉìÉNÇì«Ç›îÚÇŒÇ≥Ç»Ç¢
					//return TRUE;
				} else {
					RiffStatus = -1;
					return FALSE; 
				}
			}
*/
		}
		
		if (::memcmp(common.chunkID, "data", 4) == 0) {
			//dataÉ`ÉÉÉìÉNî≠å©
			break;
		}
		else {
			//éüÇÃÉ`ÉÉÉìÉNÇ‹Ç≈ì«Ç›îÚÇŒÇ∑
			*fSync = ::SetFilePointer(FileHandle, common.chunkSize, NULL, FILE_CURRENT);
		}
	}

	//É`ÉÉÉìÉNÉTÉCÉYÉ`ÉFÉbÉN
	wkDataSize = *dSize;
	if (common.chunkSize < wkDataSize) {
		*dSize = common.chunkSize;
	}

	*fSync += bytesRead;

	wkDataPoint = *fSync + common.chunkSize;

	return TRUE;
}


#define CHECK(x)	(::memcmp(pRMPInfo->infoID, x, 4) == 0)
#define SET(x, y)	strcpyRMP(x, (char*)pRMPInfo->data, pRMPInfo->length, y)

/*void loadRMPInfo(LIST_INFO* lplistInfo, RMPInfo* pRMPInfo)
{
	if (CHECK("INAM")) {		//track name
		SET(lplistInfo->INAM, INFO_NAME_SIZE);
	}
	else if (CHECK("IART")) {	//artist name
		SET(lplistInfo->IART, INFO_NAME_SIZE);
	}
	else if (CHECK("IPRD")) {
		SET(lplistInfo->IPRD, INFO_NAME_SIZE);
	}
	else if (CHECK("ICMT")) {
		SET(lplistInfo->ICMT, INFO_NAME_SIZE);
	}
	else if (CHECK("ICRD")) {
		SET(lplistInfo->ICRD, INFO_NAME_SIZE);
	}
	else if (CHECK("IGNR")) {
		SET(lplistInfo->IGNR, INFO_NAME_SIZE);
	}
	else if (CHECK("ICOP")) {
		SET(lplistInfo->ICOP, INFO_NAME_SIZE);
	}
	else if (CHECK("IENG")) {
		SET(lplistInfo->IENG, INFO_NAME_SIZE);
	}
	else if (CHECK("ISRC")) {
		SET(lplistInfo->ISRC, INFO_NAME_SIZE);
	}
	else if (CHECK("ISFT")) {
		SET(lplistInfo->ISFT, INFO_NAME_SIZE);
	}
	else if (CHECK("IKEY")) {
		SET(lplistInfo->IKEY, INFO_NAME_SIZE);
	}
	else if (CHECK("ITCH")) {
		SET(lplistInfo->ITCH, INFO_NAME_SIZE);
	}
	else if (CHECK("ILYC")) {
		SET(lplistInfo->ILYC, INFO_NAME_SIZE);
	}
	else if (CHECK("ICMS")) {
		SET(lplistInfo->ICMS, INFO_NAME_SIZE);
	}
}

void strcpyRMP(char* pszDest, const char* pszSrc, int length, int maxSize)
{
	if (length >= maxSize) {
		length = maxSize - 1;
		::memcpy(pszDest, pszSrc, length);
		pszDest[maxSize - 1] = 0;
	}
	else {
		::memcpy(pszDest, pszSrc, length);
		pszDest[length] = 0;
	}
}
*/