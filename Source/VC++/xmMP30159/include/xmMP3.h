// 以下の ifdef ブロックは DLL から簡単にエクスポートさせるマクロを作成する標準的な方法です。 
// この DLL 内のすべてのファイルはコマンドラインで定義された xmMP3_EXPORTS シンボル
// でコンパイルされます。このシンボルはこの DLL が使用するどのプロジェクト上でも未定義でなけ
// ればなりません。この方法ではソースファイルにこのファイルを含むすべてのプロジェクトが DLL 
// からインポートされたものとして xmMP3_API 関数を参照し、そのためこの DLL はこのマク 
// ロで定義されたシンボルをエクスポートされたものとして参照します。
#ifdef xmMP3_EXPORTS
#define xmMP3_API __declspec(dllexport)
#else
#define xmMP3_API __declspec(dllimport)
#endif

//#include <windows.h>
//#include <malloc.h>

#define INFO_NAME_SIZE		128

struct InputInfo
{
	char	szTrackName[INFO_NAME_SIZE];	//曲名
	char	szArtistName[INFO_NAME_SIZE];	//アーティスト名
	int		channels;
	int		bitRate;		//ビットレート
	int		samplingRate;	//サンプリングレート
	int		totalSec;		//総合演奏時間
};

typedef void (__stdcall *INPUT_MSG_PROC)(int msg, int param);

struct TAG_INFO
{
	char	szTrackName[INFO_NAME_SIZE];	//曲名
	char	szArtistName[INFO_NAME_SIZE];	//アーティスト名
	char    szAlbumName[INFO_NAME_SIZE];	//アルバム名
	char	szYear[5];						//リリース年号
	char	szComment[INFO_NAME_SIZE];		//コメント文字列
	int		genre;							//ジャンル
	char    szGanreName[INFO_NAME_SIZE];	//ジャンル名称
};

struct TAG_INFO_11
{
	char	szTrackName[30];	//曲名
	char	szArtistName[30];	//アーティスト名
	char    szAlbumName[30];	//アルバム名
	char	szYear[4];			//リリース年号
	char	szComment[30];		//コメント文字列
	int		genre;				//ジャンル
	char    szGanreName[128];	//ジャンル名称
	int		trackNo;			//トラック番号 -1 = v1.0
};


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

struct LYRICS_INFO
{
	int		sec;
	char	LyricsNext2[INFO_NAME_SIZE];
	char	LyricsNext1[INFO_NAME_SIZE];
	char	LyricsCurrent[INFO_NAME_SIZE];
	char	LyricsPrev1[INFO_NAME_SIZE];
	char	LyricsPrev2[INFO_NAME_SIZE];
};

struct LYRICS_INFO2
{
	long	sec;
	int		lineno;
	int		point;
	int		length;
	char	LyricsNext2[INFO_NAME_SIZE];
	char	LyricsNext1[INFO_NAME_SIZE];
	char	LyricsCurrent[INFO_NAME_SIZE];
	char	LyricsCurrentBegin[INFO_NAME_SIZE];
	char	LyricsCurrentLyrics[INFO_NAME_SIZE];
	char	LyricsCurrentAll[INFO_NAME_SIZE];
	char	LyricsPrev1[INFO_NAME_SIZE];
	char	LyricsPrev2[INFO_NAME_SIZE];
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

struct LIST_INFO
{
	char	INAM[INFO_NAME_SIZE];	//曲名
	char	IART[INFO_NAME_SIZE];	//アーティスト名
	char    IPRD[INFO_NAME_SIZE];	//製品名
	char	ICMT[INFO_NAME_SIZE];	//コメント文字列
	char	ICRD[INFO_NAME_SIZE];	//リリース年号
	char	IGNR[INFO_NAME_SIZE];	//ジャンル名
	char    ICOP[INFO_NAME_SIZE];	//著作権
	char    IENG[INFO_NAME_SIZE];	//エンジニア
	char    ISRC[INFO_NAME_SIZE];	//ソース
	char    ISFT[INFO_NAME_SIZE];	//ソフトウェア
	char    IKEY[INFO_NAME_SIZE];	//キーワード
	char    ITCH[INFO_NAME_SIZE];	//技術者
	char    ILYC[INFO_NAME_SIZE];	//歌詞
	char    ICMS[INFO_NAME_SIZE];	//コミッション
};

#define LIST_SIZE		2048

struct LIST_INFO_EX
{
	char	INAM[LIST_SIZE];	//曲名
	char	IART[LIST_SIZE];	//アーティスト名
	char    IPRD[LIST_SIZE];	//製品名
	char	ICMT[LIST_SIZE];	//コメント文字列
	char	ICRD[LIST_SIZE];	//リリース年号
	char	IGNR[LIST_SIZE];	//ジャンル名
	char    ICOP[LIST_SIZE];	//著作権
	char    IENG[LIST_SIZE];	//エンジニア
	char    ISRC[LIST_SIZE];	//ソース
	char    ISFT[LIST_SIZE];	//ソフトウェア
	char    IKEY[LIST_SIZE];	//キーワード
	char    ITCH[LIST_SIZE];	//技術者
	char    ICMS[LIST_SIZE];	//コミッション
	char	IMED[LIST_SIZE];	//中間 
	char	ISBJ[LIST_SIZE];	//タイトル
	char	IMP3[LIST_SIZE];	//MP3情報
	char    ILYC[LIST_SIZE];	//歌詞
};

#define LIST_SIZE_VB		512

struct LIST_INFO_EX_VB
{
	char	INAM[LIST_SIZE_VB];	//曲名
	char	IART[LIST_SIZE_VB];	//アーティスト名
	char    IPRD[LIST_SIZE_VB];	//製品名
	char	ICMT[LIST_SIZE_VB];	//コメント文字列
	char	ICRD[LIST_SIZE_VB];	//リリース年号
	char	IGNR[LIST_SIZE_VB];	//ジャンル名
	char    ICOP[LIST_SIZE_VB];	//著作権
	char    IENG[LIST_SIZE_VB];	//エンジニア
	char    ISRC[LIST_SIZE_VB];	//ソース
	char    ISFT[LIST_SIZE_VB];	//ソフトウェア
	char    IKEY[LIST_SIZE_VB];	//キーワード
	char    ITCH[LIST_SIZE_VB];	//技術者
	char    ICMS[LIST_SIZE_VB];	//コミッション
	char	IMED[LIST_SIZE_VB];	//中間 
	char	ISBJ[LIST_SIZE_VB];	//タイトル
	char	IMP3[LIST_SIZE_VB];	//MP3情報
	char    ILYC[LIST_SIZE_VB];	//歌詞
};

struct WAVE_DATA
{
	int		channels;
	int		bitsPerSample;
	int		left;
	int		right;
};

struct WAVE_FORM
{
	int		channels;			//チャネル数
	int		bitsPerSample;		//ビット数/1サンプル
	int		samplingRate;
	int		dataSize;			//ファイルサイズ
};
 
struct LYRICS3_INFO
{
	char	IND_LYR;		//LYRフィールドの有無 '0':無し '1':有り
	char	IND_TIMETAG;	//タイムタグの有無 '0':無し '1':有り
	char	LYR[60000];		//歌詞フィールド
	char	INF[60000];		//情報フィールド
	char	AUT[250];		//作詞／作曲者名
	char	EAL[250];		//拡張アルバム名
	char	EAR[250];		//拡張アーティスト名
	char	ETT[250];		//拡張トラックタイトル
	char	IMG[60000];		//イメージファイルへのリンク
};

struct LYRICS3_INFO_VB
{
	char	IND_LYR;		//LYRフィールドの有無 '0':無し '1':有り
	char	IND_TIMETAG;	//タイムタグの有無 '0':無し '1':有り
	char	AUT[250];		//作詞／作曲者名
	char	EAL[250];		//拡張アルバム名
	char	EAR[250];		//拡張アーティスト名
	char	ETT[250];		//拡張トラックタイトル
};

//xmMP3

//エンコード系
xmMP3_API int __stdcall  xmMP3_getEncodeState(int* readSize, int* encodeSize);
xmMP3_API BOOL __stdcall xmMP3_encodeOpen(const char* pszWaveName, WAVE_FORM* pWaveForm);
xmMP3_API BOOL __stdcall xmMP3_encodeStart(const char* pszMp3Name);
xmMP3_API BOOL __stdcall xmMP3_encodeStop();

//未使用
xmMP3_API void __stdcall xmMP3_startAnalyze();
xmMP3_API void __stdcall xmMP3_stopAnalyze();
xmMP3_API void __stdcall xmMP3_startAnalyzeThread();
xmMP3_API void __stdcall xmMP3_stopAnalyzeThread();
xmMP3_API BOOL __stdcall xmMP3_getWaveData(WAVE_DATA* pWaveData);

//情報系
xmMP3_API BOOL __stdcall xmMP3_setLyricsFile(const char* pszLyricsName);
xmMP3_API BOOL __stdcall xmMP3_getLyrics(LYRICS_INFO* pLyricsInfo);
xmMP3_API void __stdcall xmMP3_getSpectrum(int* pSpecL, int* pSpecR);
xmMP3_API void __stdcall xmMP3_getWave(int* pWaveL, int* pWaveR);
xmMP3_API BOOL __stdcall xmMP3_getFileLyrics3Info(const char* pszName, LYRICS3_INFO* pLyrics3Info);
xmMP3_API BOOL __stdcall xmMP3_getFileLyrics3InfoVB(const char* pszName, LYRICS3_INFO_VB* pLyrics3InfoVb, char* lyrData, char* InfData, char* ImgData);
xmMP3_API void __stdcall xmMP3_setLyrics3Use(int useLyrics3);
xmMP3_API BOOL __stdcall xmMP3_readLyrics3Data();
xmMP3_API BOOL __stdcall xmMP3_clearLyrics();
xmMP3_API void __stdcall xmMP3_setKaraokeUse(int useKaraoke);
xmMP3_API BOOL __stdcall xmMP3_getLyrics2(LYRICS_INFO2* pLyricsInfo);
xmMP3_API void __stdcall xmMP3_setLyricsTime(int flag);

//コールバック関数
xmMP3_API BOOL __stdcall xmMP3_callback(INPUT_MSG_PROC pProc);
xmMP3_API BOOL __stdcall xmMP3_startCallback();
xmMP3_API BOOL __stdcall xmMP3_stopCallback();

//ファイル出力系
xmMP3_API BOOL __stdcall xmMP3_delFileLyrics3Info(const char* pszName);
xmMP3_API void __stdcall xmMP3_setLyrics3InsField(int field);
xmMP3_API BOOL __stdcall xmMP3_setFileLyrics3Info(const char* pszName, LYRICS3_INFO* pLyrics3Info);
xmMP3_API BOOL __stdcall xmMP3_setFileLyrics3InfoVB(const char* pszName, LYRICS3_INFO_VB* pLyrics3InfoVb, char* lyrData, char* InfData, char* ImgData);
xmMP3_API BOOL __stdcall xmMP3_playDecodeWave(const char* pszWaveName);
xmMP3_API BOOL __stdcall xmMP3_decodeWave(const char* pszWaveName);
xmMP3_API BOOL __stdcall xmMP3_changeWav(const char* pszName);
xmMP3_API BOOL __stdcall xmMP3_changeMp3(const char* pszName);
xmMP3_API BOOL __stdcall xmMP3_changeRmp(const char* pszName);
xmMP3_API BOOL __stdcall xmMP3_cutMacBinary(const char* pszName);
xmMP3_API BOOL __stdcall xmMP3_setTagInfo(const char* pszName, TAG_INFO* pTagInfo,int tagSet, int tagAdd);
xmMP3_API BOOL __stdcall xmMP3_setListInfo(const char* pszName, LIST_INFO* pListInfo);
xmMP3_API BOOL __stdcall xmMP3_setListInfoEX(const char* pszName, LIST_INFO_EX* pListInfo);
xmMP3_API BOOL __stdcall xmMP3_setListInfoExVB(const char* pszName, LIST_INFO_EX_VB* pListInfo);
xmMP3_API BOOL __stdcall xmMP3_setTagInfoEX(const char* pszName, TAG_INFO_11* pTagInfo,int tagSet, int tagAdd);

//情報取得系
xmMP3_API BOOL __stdcall xmMP3_getFileInfoMtr(const char* pszName, TAG_INFO* pTagInfo, MPEG_INFO* pMpegInfo, LIST_INFO* pListInfo, int* pFileType);
xmMP3_API BOOL __stdcall xmMP3_getFileInfo2(const char* pszName, TAG_INFO* pTagInfo, MPEG_INFO* pMpegInfo, LIST_INFO* pListInfo);
xmMP3_API BOOL __stdcall xmMP3_getFileInfo(const char* pszName, TAG_INFO* pTagInfo, MPEG_INFO* pMpegInfo);
xmMP3_API BOOL __stdcall xmMP3_getFileTagInfo(const char* pszName, TAG_INFO* pTagInfo);
xmMP3_API BOOL __stdcall xmMP3_getListInfo(LIST_INFO* pListInfo);
xmMP3_API BOOL __stdcall xmMP3_getMpegInfo(MPEG_INFO* pMpegInfo);
xmMP3_API BOOL __stdcall xmMP3_getTagInfo(TAG_INFO* pTagInfo);
xmMP3_API int __stdcall xmMP3_getFileType(const char* pszName);
xmMP3_API BOOL __stdcall xmMP3_getGenre(TAG_INFO* pTagInfo);
xmMP3_API long __stdcall xmMP3_getWinampPlayMs();
xmMP3_API int __stdcall xmMP3_getWinampTotalSec();
xmMP3_API int __stdcall xmMP3_getPlayBitRate();
xmMP3_API int __stdcall xmMP3_getLastErrorNo();
xmMP3_API int __stdcall xmMP3_getWaveOutSupport();
xmMP3_API int __stdcall xmMP3_getSilentFrames(const char* pszName);
xmMP3_API BOOL __stdcall xmMP3_getListInfoEX(const char* pszName, LIST_INFO_EX* pListInfo);
xmMP3_API BOOL __stdcall xmMP3_getListInfoExVB(const char* pszName, LIST_INFO_EX_VB* pListInfo);
xmMP3_API BOOL __stdcall xmMP3_getFileTagInfoEX(const char* pszName, TAG_INFO_11* pTagInfo);

//xmMP3 関連
xmMP3_API int __stdcall xmMP3_getVersion();
xmMP3_API BOOL __stdcall xmMP3_setxmMP3Option(xmMP3_OPTION* pxmMP3Option);
xmMP3_API void __stdcall xmMP3_getxmMP3Option(xmMP3_OPTION* pxmMP3Option);
xmMP3_API BOOL __stdcall xmMP3_setDecodeOption(DEC_OPTION* pDecOption);
xmMP3_API void __stdcall xmMP3_getDecodeOption(DEC_OPTION* pDecOption);
xmMP3_API long __stdcall xmMP3_debug();
xmMP3_API void __stdcall xmMP3_setEqualizer(int* pTable);
xmMP3_API void __stdcall xmMP3_setFftWindow(int window);
xmMP3_API void __stdcall xmMP3_setWaveOutDeviceId(int id);

//基本操作系
xmMP3_API BOOL __stdcall xmMP3_setStepPitch(int pitch, int frames);
xmMP3_API int __stdcall xmMP3_getStepPitch();
xmMP3_API BOOL __stdcall xmMP3_reload();
xmMP3_API BOOL __stdcall xmMP3_setPitch(int pitch);
xmMP3_API int __stdcall xmMP3_getPitch();
xmMP3_API BOOL __stdcall xmMP3_init(int flag);
xmMP3_API BOOL __stdcall xmMP3_free();
xmMP3_API BOOL __stdcall xmMP3_open(const char* pszName, InputInfo* pInfo);
xmMP3_API BOOL __stdcall xmMP3_close();
xmMP3_API int __stdcall xmMP3_getState(int* sec);
xmMP3_API BOOL __stdcall xmMP3_play();
xmMP3_API BOOL __stdcall xmMP3_stop();
xmMP3_API BOOL __stdcall xmMP3_pause();
xmMP3_API BOOL __stdcall xmMP3_restart();
xmMP3_API BOOL __stdcall xmMP3_seek(int seekSec);
xmMP3_API int __stdcall xmMP3_getPlayFlames();
xmMP3_API BOOL __stdcall xmMP3_setPlayFlames(int flames);
xmMP3_API long __stdcall xmMP3_getPlaySamples();
xmMP3_API int __stdcall xmMP3_getTotalSamples();
xmMP3_API BOOL __stdcall xmMP3_setPlaySamples(int sample);
xmMP3_API BOOL __stdcall xmMP3_setVolume(int left, int right);
xmMP3_API BOOL __stdcall xmMP3_getVolume(int* left, int* right);
xmMP3_API void __stdcall xmMP3_setFadeIn(int on);
xmMP3_API void __stdcall xmMP3_setFadeOut(int on);
xmMP3_API void __stdcall xmMP3_fadeOut();
xmMP3_API void __stdcall xmMP3_setOverTime(int on);
xmMP3_API void __stdcall xmMP3_setSoftVolume(int left, int right);
xmMP3_API void __stdcall xmMP3_getSoftVolume(int* left, int* right);


//MP3Input

void input_setWaveOutDeviceId(int id);

void input_setLyricsTime(int flag);

void input_setFftWindow(int window);
BOOL input_getListInfoEX(const char* pszName, LIST_INFO_EX* lpListInfo);
BOOL input_setListInfoEX(const char* pszName, LIST_INFO_EX* lpListInfo);
BOOL getFileTagInfoEX(const char* pszName, TAG_INFO_11* pTagInfo);
BOOL setTagInfoEX(const char* pszName, TAG_INFO_11* pTagInfo,int tagSet, int tagAdd);
void input_setKaraokeUse(int useKaraoke);
BOOL input_getLyrics2(LYRICS_INFO2* pLyricsInfo);

BOOL input_clearLyrics();
BOOL input_readLyrics3Data();
void input_setLyrics3Use(int useLyrics3);
BOOL input_getFileLyrics3Info(const char* pszName, LYRICS3_INFO* lpLyrics3Info);
BOOL input_setFileLyrics3Info(const char* pszName, LYRICS3_INFO* lpLyrics3Info, int setFlag);

void input_setLyrics3InsField(int field);

BOOL input_playDecodeWave(const char* pszWaveName);

int input_getSilentFrames(const char* pszName);

int input_getWaveOutSupport();

void input_setSoftVolume(int left, int right);
void input_getSoftVolume(int* left, int* right);

void input_setOverTime(int on);

int input_getLastErrorNo();

void input_setFadeIn(int on);
void input_setFadeOut(int on);
void input_fadeOut();

BOOL input_reload();

BOOL input_setStepPitch(int pitch, int frames);
int input_getStepPitch();
BOOL input_setPitch(int pitch);
int input_getPitch();

int input_getPlayFlames();
DWORD input_getPlaySamples();
int input_getTotalSamples();

/*
void input_startAnalyze();
void input_stopAnalyze();
void input_startAnalyzeThread();
void input_stopAnalyzeThread();
*/

BOOL input_decodeWave(const char* pszWaveName);
long input_getWinampPlayMs();
int input_getWinampTotalSec();
int input_getPlayBitRate();

int getFileType(const char* pszName);
BOOL input_cutMacBinary(const char* pszName);
BOOL input_changeMp3(const char* pszName);
BOOL input_changeWav(const char* pszName);
BOOL input_changeRmp(const char* pszName);
BOOL setListInfo(const char* pszName, LIST_INFO* pListInfo);

BOOL getListInfo(LIST_INFO* pListInfo);
BOOL getFileTagInfo(const char* pszName, TAG_INFO* pTagInfo);
BOOL getFileInfo(const char* pszName, TAG_INFO* pTagInfo, MPEG_INFO* pMpegInfo);
BOOL getFileInfo2(const char* pszName, TAG_INFO* pTagInfo, MPEG_INFO* pMpegInfo, LIST_INFO* pListInfo);
BOOL getFileInfoMtr(const char* pszName, TAG_INFO* pTagInfo, MPEG_INFO* pMpegInfo, LIST_INFO* pListInfo,int* pFileType);
long input_debug();
BOOL input_GetGenre(TAG_INFO* pTagInfo);
BOOL input_SetxmMP3Option(xmMP3_OPTION* pxmMP3Option);
void input_GetxmMP3Option(xmMP3_OPTION* pxmMP3Option);
BOOL input_SetDecodeOption(DEC_OPTION* pDecOption);
void input_GetDecodeOption(DEC_OPTION* pDecOption);
BOOL getTagInfo(TAG_INFO* pTagInfo);
BOOL setTagInfo(const char* pszName, TAG_INFO* pTagInfo,int tagSet, int tagAdd);
BOOL getMpegInfo(MPEG_INFO* pMpegInfo);
int getTime();
BOOL initDec();
BOOL freeDec();

BOOL input_setCallback(INPUT_MSG_PROC pProc);
BOOL input_startCallback();
BOOL input_stopCallback();

BOOL input_open(const char* pszName, InputInfo* pInfo);
BOOL input_close();
int input_getState();
//X-MaD
BOOL output_SetBufferRedim(int BufferRedimSize);
BOOL input_SetReadHaedA(int NEqualFrame,int MaxHeadSize);
void output_SetXSound(int mode,int option);
int output_GetXSound();
double output_getXNormLevel();
double output_setXNormLevel(double level);
//X-MaD
BOOL input_play();
BOOL input_stop();
BOOL input_pause();
BOOL input_restart();
void input_seekBySec(int sec);
void input_seekByFrame(int frame);
void input_seekBySample(int sample);
BOOL input_setVolume(int left, int right);
BOOL input_getVolume(int* left, int* right);
void input_setEqualizer(int* pTable);

void input_getSpectrum(int* pSpecL, int* pSpecR);
void input_getWave(int* pWaveL, int* pWaveR);

BOOL input_setLyricsFile(const char* pszLyricsName);
BOOL input_getLyrics(LYRICS_INFO* pLyricsInfo);

//MP3Enc
int  enc_getEncodeState(int* readSize, int* encodeSize);
BOOL enc_open(const char* pszWaveName, WAVE_FORM* pWaveForm);
BOOL enc_start(const char* pszMp3Name);
BOOL enc_stop();







