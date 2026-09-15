#include <windows.h>
#include <process.h>	//_beginthreadex, _endthreadex
#include <string.h>		//memmove
#include <stdio.h>		//fgets
#include "..\\..\\mp3dec\\include\\mp3dec.h"

//ライブラリの追加
//#pragma comment(lib, "winmm.lib")
//#ifdef _DEBUG
//#pragma comment(lib, "mpegdecodeD.lib")
//#else
//#pragma comment(lib, "mpegdecode.lib")
//#endif

//lib版作成時にはコメントにすること
//#pragma comment(lib, "..\\mp3dec\\Release\\mp3dec.lib")


#pragma pack(1)
struct RiffChunk {
	BYTE	chunkID[4];		//'RIFF'
	DWORD	chunkSize;
	BYTE	formType[4];	//'WAVE','RMP3'
};

struct CommonChunk {
	BYTE	chunkID[4];
	DWORD	chunkSize;	
};

/*struct RMPInfo {
	BYTE	infoID[4];
	DWORD	length;
	BYTE	data[1];
};*/
#pragma pack()

#define INFO_NAME_SIZE		128

struct InputInfo
{
	char	szTrackName[INFO_NAME_SIZE];	//曲名
	char	szArtistName[INFO_NAME_SIZE];	//アーティスト名
	int		channels;		//チャンネル数
	int		bitRate;		//ビットレート
	int		samplingRate;	//サンプリングレート
	int		totalSec;		//総演奏時間
};

typedef void (__stdcall *INPUT_MSG_PROC)(int msg, int param);

struct MPEG_INFO
{
	int		version;		//MPEGのバージョン
	int		layer;			//MPEGのレイヤ
	int		crcDisable;		//エラー保護
	int		extension;		//個人情報
	int		mode;			//チャンネルモード
	int		copyright;		//著作権
	int		original;		//オリジナル
	int		emphasis;		//エンフォシス		

	int		channels;		//出力チャンネル数
	int		bitRate;		//ビットレート
	int		samplingRate;	//サンプリングレート
	int		fileSize;		//ファイルサイズ
	int		flames;			//フレーム数
	int		totalSec;		//総合演奏時間
};

struct DEC_OPTION{
	int		reduction;		//サンプリング 0:1/1 1:1/2 2:1/4 (Default = 0)
	int		convert;		//チャンネル 0:ステレオ 1:モノラル(Default = 0)
	int		freqLimit;		//周波数
};

struct xmMP3_OPTION
{
	int inputBlock;			//入力フレーム数
	int outputBlock;		//出力フレーム数
	int inputSleep;			//入力直後のスリープ時間(ミリ秒)
	int outputSleep;		//出力直後のスリープ時間(ミリ秒)
};

struct WAVE_DATA
{
	int		channels;
	int		bitsPerSample;
	int		left;
	int		right;
};

/* 新定義 */
HANDLE			m_hFile;					//対象のファイル
MPEG_DECODE_INFO m_decInfo;
HANDLE			m_hThread;					//デコード用スレッド
int				m_threadID;					//デコード用スレッドID
int				m_state;					//現在の状態
//X-MaD
int             m_NEqualFrame = 128;
int				m_MaxHeadSize = 102400;

INPUT_MSG_PROC	m_pProc;					//コールバック関数
int				m_userData;					//コールバック関数パラメータ

int				m_playSec;					//演奏経過時間
int				m_playSecCur;				//演奏経過時間(出力用)
int				m_totalSec;					//演奏総時間
double			m_playMs;					//演奏中の時間（ミリ秒）
int				m_playFrames;				//演奏フレーム
DWORD			m_playSamples;				//演奏サンプル
double			m_playSpeed;
int				m_reduction;
int				m_bitRate;

int				m_firstSync;				//最初の同期ヘッダ位置
int				m_firstSample;				//サンプルの位置
double			m_msPerFrame;				//1フレームの演奏時間（ミリ秒）

int				m_inputBlock = 30;			//入力フレーム数
int				m_outputBlock_in = 20;		//出力フレーム数
int				m_outputBlock_in_tmp = 20;		//出力フレーム数
int				m_inputSleep = 5;			//入力直後のスリープ時間(ミリ秒)

BYTE*			m_inputBuf;					//入力バッファ（MPEG Audio）
BYTE*			m_inputPtr;					//入力ポインタ
int				m_inputSize;				//入力バッファサイズ
int				m_inputBytes;				//デコードしていないバッファサイズ
BYTE*			m_outputBuf_in;				//出力バッファ（PCM）

int				m_lvol;
int				m_rvol;
int				m_olvol;
int				m_orvol;
BOOL			m_fadeOutFlag = false;

/* 旧定義 */
MPEG_DECODE_OPTION   decodeOption;			//デコードエンジンオプション
DWORD			dataSize;					//MPEG Audioデータサイズ
int				fileSize;					//MPEG ファイルサイズ
int				decodePos;					//再生時間計算用

BOOL			callBackFlag = FALSE;		//コールバック関数使用有無
int				decFlag;					//デコードフラグ 0:再生 1:Wave出力 2:再生しながらWAVE出力
HANDLE			m_hWaveFile;				//WAVEファイル

int				m_FileType;					//ﾌｧｲﾙ種別

int				m_winampTotalSec;
long            debugNo;					//デバッグ用

int				m_platForm;						//1:WinNT 0:その他

int				m_readWaveDataSize;

//int				m_bitsPerSample;

//WAVE_DATA		m_wave_data;

//#define DEFAULT_BUFFER_SIZE	44100
#define DEFAULT_BUFFER_SIZE	4096
int				m_outputBlock_in_wav = 5;		//出力フレーム数

int            RiffStatus;					//MpegCheck
WAVEFORMATEX	m_waveFormat;				//WAVEフォーマット
BOOL			m_waveFlag;					//オープンファイルがPCMならTRUE
int				m_outWaveBufSize;
int				m_inWaveBufSize;

BOOL			m_openFlag = FALSE;
BOOL			m_playFlag = FALSE;
BOOL			m_timeFlag = TRUE;			//TRUE:演奏時間が総時間越したら切り捨て

BOOL			m_seekFlag = FALSE;			//シーク中なら TRUE
BOOL			m_stepFlag = FALSE;

#define ALLOW_DATA_SIZE		268435455

//デコードスレッド用メッセージ
enum {
	EV_STOP		= (WM_APP + 0),
	EV_PAUSE	= (WM_APP + 2),
	EV_RESTART	= (WM_APP + 3),
	EV_SEEK		= (WM_APP + 4),
	EV_EQUALIZER= (WM_APP + 5),
};

//現在の状態
//getState関数で取得できる
enum {
	STATE_STOP	= 0,
	STATE_PLAY	= 1,
	STATE_PAUSE	= 2,
	STATE_SEEK	= 3,
};

int				m_errNo;

//エラーステータス
enum {
	ERR_MP3_FILE_OPEN	= 1,
	ERR_MP3_FILE_NOT_OPEN	= 2,
	ERR_MP3_FILE_READ	= 3,
	ERR_MP3_FILE_WRITE	= 4,
	ERR_WAV_FILE_OPEN	= 5,
	ERR_WAV_FORMAT	= 6,
	ERR_ENCODE_FILE_OPEN	= 7,
	ERR_LYRICS_FILE_OPEN	= 8,
	ERR_LYRICS_NON_DATA	= 9,
	ERR_FRAME_HEADER_NOT_FOUND = 10,
	ERR_FRAME_HEADER_READ = 11,
	ERR_STATE_STOP	= 12,
	ERR_NOT_STATE_STOP	= 13,
	ERR_NOT_STATE_PLAY	= 14,
	ERR_STATE_NON_ENCODE = 15,
	ERR_PLAY = 16,
	ERR_STOP = 17,
	ERR_INVALID_VALUE	= 18,
	ERR_MALLOC	= 19,
	ERR_NON_RIFF = 20,
	ERR_RIFF = 21,
	ERR_NOT_MP3 = 22,
	ERR_MAC_BIN = 23,
	ERR_UNKNOWN_FILE = 24,
	ERR_OPEN_OUT_DEVICE = 25,
	ERR_DECODE = 26,
	ERR_DECODE_THREAD = 27,
	ERR_ENCODE_THREAD = 28,
	ERR_CREATE_EVENT = 29,
	ERR_CODEC_NOT_FOUND = 30,
	ERR_WAVE_TABLE_NOT_FOUND = 31,
	ERR_ACM_OPEN = 32,
	ERR_PLATFORM_NT = 33,
};

//メッセージ番号
//setCallback関数で指定した関数に渡される
/*
enum {
	MSG_ERROR		= 0,
	MSG_PLAYDONE	= 1,
	MSG_PLAYING		= 2,
	MSG_PAUSING		= 3,
	MSG_STOPING		= 4,
};
*/
#define INPUT_MSG_ERROR			0
#define INPUT_MSG_STOP			1
#define INPUT_MSG_PLAY			2
#define INPUT_MSG_PAUSE			3
#define INPUT_MSG_PLAYDONE		4
#define INPUT_MSG_BITRATE		5
#define INPUT_MSG_WAVE_LEFT		6
#define INPUT_MSG_WAVE_RIGHT	7

static const double ms_table[3][3] =
{
	// Layer1
	{8.707483f, 8.0f, 12.0f},
	// Layer2
	{26.12245f, 24.0f, 36.0f},
	// Layer3
	{26.12245f, 24.0f, 36.0f},
};


//MP3Input

void input_setWaveOutDeviceId(int id);

void input_setFftWindow(int window);
BOOL input_playDecodeWave(const char* pszWaveName);

int input_getSilentFrames(const char* pszName);

int input_getWaveOutSupport();

void input_setSoftVolume(int left, int right);
void input_getSoftVolume(int* left, int* right);

void input_setOverTime(int on);

void input_setErrNo(int no);

int input_getDecodeOutputSize();
int input_getTotalOutputSize();

void input_fadeOut();
void input_setFadeIn(int on);
void input_setFadeOut(int on);
void input_fadeOutVolume();

BOOL input_reload();

BOOL input_setStepPitch(int pitch, int frames);
int input_getStepPitch();
BOOL input_setPitch(int pitch);
int input_getPitch();

void input_callBackTime();

int input_getPlayFlames();
DWORD input_getPlaySamples();
int input_getTotalSamples();

BOOL input_decodeWave(const char* pszWaveName);
long input_getWinampPlayMs();
int input_getWinampTotalSec();
int input_getPlayBitRate();

long input_debug();
void input_setDebug(long no);
BOOL input_SetxmMP3Option(xmMP3_OPTION* pxmMP3Option);
void input_GetxmMP3Option(xmMP3_OPTION* pxmMP3Option);
BOOL input_SetDecodeOption(DEC_OPTION* pDecOption);
void input_GetDecodeOption(DEC_OPTION* pDecOption);
BOOL getMpegInfo(MPEG_INFO* pMpegInfo);
int getTime();

BOOL initDec();
BOOL freeDec();

BOOL input_setCallback(INPUT_MSG_PROC pProc);
BOOL input_startCallback();
BOOL input_stopCallback();

BOOL input_open(const char* pszName, InputInfo* pInfo);
BOOL input_close();
int input_read(void* pData, int length);
int input_getState();
BOOL input_play();
BOOL input_stop();
BOOL input_pause();
BOOL input_restart();
//BOOL input_seek(int pos);

void input_seekBySec(int sec);
void input_seekByFrame(int frame);
void input_seekBySample(int sample);
void input_setEqualizer(int* pTable);

int input_FrameSizeByte(void);
int input_bytesToSample(int bytes);
int input_secToSample(int sec);
int input_sampleToSec(int sample);
long input_sampleToMs(int sample);
BOOL input_skipNextSync();

BOOL input_setVolume(int left, int right);
BOOL input_getVolume(int* left, int* right);

void input_getSpectrum(int* pSpecL, int* pSpecR);
void input_getWave(int* pWaveL, int* pWaveR);

BOOL input_sendEvent(int code, int param);
int input_ThreadProc();
int input_decode();
void input_handleEvent(const MSG& msg);
void input_actionStop(BOOL wait);
void input_actionPause();
void input_actionRestart();
void input_actionSeek(int sample);

//RIFF
BOOL readRIFF(HANDLE FileHandle, int* fSync, DWORD* dSize);
//#define NUM_WAVEHDR			4				//データブロック数

//MP3Output
struct OutputInfo
{
	int		channels;			//チャネル数
	int		bitsPerSample;		//ビット数/1サンプル
	int		frequency;
	int		bufferSize;
};

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
BOOL output_open(const char* pszName, OutputInfo& info);
BOOL output_close();
int output_write(const void* pData, int length);
BOOL output_stop();
BOOL output_pause();
BOOL output_restart();
BOOL output_setVolume(int left, int right);
BOOL output_getVolume(int* left, int* right);
DWORD output_getCurrentSample();
int input_GetState();
int output_getWave(DWORD sample, int* pWaveL, int* pWaveR);
int output_getSpectrum(DWORD sample, int* pSpecL, int* pSpecR);
int output_sampleToBytes(DWORD sample);

void output_setFadeIn(int on);
void output_setFadeOut(int on);
void output_softVolumeMax();

//MP3Enc
void initEnc();
void freeEnc();

void input_actionSpeedSeek(int seekSample);
