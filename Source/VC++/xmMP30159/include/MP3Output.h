#include <windows.h>
#include <process.h>
#include <stdio.h>

#pragma comment(lib, "winmm.lib")

class Exception
{
public:
	int		code;

	Exception(int c) : code(c) {}
};

class MMException
{
public:
	int		code;

	MMException(int c) : code(c) {}
};


struct OutputInfo
{
	int		channels;			//チャネル数
	int		bitsPerSample;		//ビット数/1サンプル
	int		frequency;
	int		bufferSize;
};

struct xmMP3_OPTION
{
	int inputBlock;			//入力フレーム数
	int outputBlock;		//出力フレーム数
	int inputSleep;			//入力直後のスリープ時間(ミリ秒)
	int outputSleep;		//出力直後のスリープ時間(ミリ秒)
};

HWAVEOUT		m_hWaveOut;					//オーディオ出力デバイス
//HANDLE			m_out_hThread;
//int				m_out_threadID;

HANDLE			m_hSemaphore;				//データブロック用セマフォ
CRITICAL_SECTION m_cs;						//各種変数 同期用

int				m_channels = 0;
int				m_bitsPerSample;
int				m_xsoundmode = 0;
int				m_xsoundoption = 0;
double			m_xnorm = 1;
double			m_xnormset = 0;
int				m_blockSize = 1;
int				m_BufferRedim = 1;

int				m_outputBlock;
int				m_outputIndex;
int				m_outputFinished;
BYTE**			m_outputTable;

//boolean			m_analyze = false;
//boolean			m_analyzeThread = false;
//int				m_specTableL[256];
//int				m_specTableR[256];

int				m_waveBlock;
int				m_waveIndex;				//WAVEHDR インデックス
//int				m_waveStack;				//未再生 WAVEHDR 数
//int				m_waveCount = 16;				//WAVEHDR 配列数
WAVEHDR*		m_waveTable;				//WAVEHDR 格納テーブル

BYTE*			m_outputPtr;
int				m_outputBytes;
int				m_outputSize;
int				m_pitch = 100;				//ピッチ(%)

/*
BYTE*		m_outputBuf;
BYTE*		m_outputPtr;
int			m_outputSize;
int			m_outputBytes;
*/
int				m_outputSleep = 0;

int				m_lSoftVol = 100;						//ソフトウェアボリューム(左)
int				m_rSoftVol = 100;						//ソフトウェアボリューム(右)
int				m_lMaxSoftVol = 100;			//最大ソフトウェアボリューム(左)
int				m_rMaxSoftVol = 100;			//最大ソフトウェアボリューム(右)
BOOL			m_fadeIn = false;
BOOL			m_fadeOut= false;

BOOL			m_stepPitch;

int				m_deviceId = WAVE_MAPPER;		//デバイスID

enum {
	rectangle	= 0,
	hanning		= 1,
	hamming		= 2,
	blackman	= 3,
};

int				m_fftWindow = hanning;

//現在の状態
//getState関数で取得できる
enum {
	STATE_STOP	= 0,
	STATE_PLAY	= 1,
	STATE_PAUSE	= 2,
	STATE_SEEK	= 3,
};
//

#define ALLOW_DATA_SIZE		268435455

void output_setWaveOutDeviceId(int id);

void output_setFftWindow(int window);

int output_getWaveOutSupport();

void output_setSoftVolume(int left, int right);
void output_getSoftVolume(int* left, int* right);

void output_setStepPitch(int pitch);
int output_getStepPitch();
void output_setPitch(int pitch);
int output_getPitch();

BOOL output_SetxmMP3Option(xmMP3_OPTION* pxmMP3Option);
void output_GetxmMP3Option(xmMP3_OPTION* pxmMP3Option);

BOOL initOutput();
BOOL freeOutput();
void output_xsound( short *buf, int size );
void output_xnormalize( short *buf, int size );
void output_applynormalize( short *buf, int size, double normalize );
double output_calcnormalize( short *buf, int size);
//BOOL SetBufferRedim(int BufferRedimSize);
BOOL output_open(const char* pszName, OutputInfo& info);
BOOL output_close();
int output_write(const void* pData, int length);
BOOL output_stop();
BOOL output_pause();
BOOL output_restart();
BOOL output_setVolume(int left, int right);
BOOL output_getVolume(int* left, int* right);

//void output_getSpectrum(int* pSpecL, int* pSpecR);
DWORD output_getCurrentSample();
int output_getWave(int sample, int* pWaveL, int* pWaveR);
int output_getSpectrum(DWORD sample, int* pSpecL, int* pSpecR);

int output_bytesToSample(int bytes);
int output_sampleToBytes(DWORD sample);
void output_playBlock();

void CALLBACK WaveProc(HWAVEOUT hWaveOut, UINT msg,
		DWORD instance, DWORD param1, DWORD param2);

//WAVE操作
void output_softVolumeMax();
void output_chgSoftVol( short *buf, int size );
void output_chgSoftVol8( unsigned char *buf, int size );
void output_setFadeIn(int on);
void output_setFadeOut(int on);
void output_chgStepSpeed( short *buf, int size );

//inputMP3
void input_callBackTime();
int input_getDecodeOutputSize();
int input_getTotalOutputSize();
void input_fadeOutVolume();
void input_setDebug(long no);
void input_seekBySample(int sample);
int input_GetState();

void fftCopy(int pos, short s);
void fftTransform(int* spec);
