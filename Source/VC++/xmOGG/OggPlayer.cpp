#include "stdafx.h"
#include "oggplayer.h"
#define DS_QUAD_A      0.000002131
#define DS_QUAD_B     -0.00058625
#define DS_QUAD_C      0.058645625
#define DS_QUAD_D     -2.73303125
#define DS_QUAD_E   60.0

#include <strstream>
using namespace std;
DWORD shorty=0;

FILE		*stream;
int BUFSIZE=65536*10;

int IsGlobal=0;

void openstream() {
	stream=fopen("debugstream.txt","a+");
}

void closestream() {
	fclose(stream);
}

OggPlayer::OggPlayer()
{
    bFileOpened     = false;
    bInitialized    = false;
    bReleaseDS      = false;
    pDS             = NULL;
    pDSB            = NULL;
    bLoop           = false;
    bDone           = false;
    bAlmostDone     = false;
}

OggPlayer::~OggPlayer()
{
    if (bFileOpened)
        Close();

    if (bReleaseDS && pDS)
        pDS->Release();
}

bool OggPlayer::InitDirectSound( HWND hWnd, int global )
{
	HRESULT hr;
	IsGlobal=global;
	
	if (FAILED(hr = DirectSoundCreate(NULL, &pDS, NULL)))
		return bInitialized = false;

	pDS->Initialize(NULL);
	if (IsGlobal==1) {
		pDS->SetCooperativeLevel( hWnd, DSSCL_EXCLUSIVE ); //DSSCL_NORMAL ); // DSSCL_PRIORITY );
	} else {
		pDS->SetCooperativeLevel( hWnd, DSSCL_NORMAL );
	}

    bReleaseDS = true;

    return bInitialized = true;
}

void OggPlayer::SetPan( int Position ) {
	// expecting range between -100 and 100
	if (Position<-100) Position=-100;
	if (Position>100) Position=100;

	long actual=(long)Position*100;

	double result = pDSB->SetPan(actual);
}

void OggPlayer::SetVolume( int Volume ) {
	if (Volume>100) Volume=100;
	if (Volume<0) Volume=0;

	double percent=(double)Volume;
	long actual=(LONG) (-100.0 *
		((DS_QUAD_A * percent * percent *
		percent * percent) +
		(DS_QUAD_B * percent * percent *
		percent) +
       (DS_QUAD_C * percent * percent) +
       (DS_QUAD_D * percent) +
        DS_QUAD_E));
	
	double result = pDSB->SetVolume(actual);
	/*if (result != DS_OK){
		openstream();
		fprintf(stream,"DirectSound Error: %d\nPossible errors are "
			"CONTROLUNAVAIL (%d), GENERIC (%d), INVALIDPARAM (%d), "
			"and PRIOLEVELNEEDED (%d)",result,DSERR_CONTROLUNAVAIL,DSERR_GENERIC,DSERR_INVALIDPARAM,DSERR_PRIOLEVELNEEDED);
		closestream();
	}*/
}

bool OggPlayer::OpenOgg( char *filename )
{
	OggPlayer::Close();
	OggPlayer::OggPlayer();
	
	if (!bInitialized)
        return false;

    if (bFileOpened)
        Close();

	FILE    *f;

    f = fopen(filename, "rb");
    if (!f) return false;

    ov_open(f, &vf, NULL, 0);

    // ok now the tricky part

    // the vorbis_info struct keeps the most of the interesting format info
    vorbis_info *vi = ov_info(&vf,-1);

    // set the wave format
	WAVEFORMATEX	    wfm;

    memset(&wfm, 0, sizeof(wfm));

    wfm.cbSize          = sizeof(wfm);
    wfm.nChannels       = vi->channels;
    wfm.wBitsPerSample  = 16;                    // ogg vorbis is always 16 bit
    wfm.nSamplesPerSec  = vi->rate;
    wfm.nAvgBytesPerSec = wfm.nSamplesPerSec*wfm.nChannels*2;
    wfm.nBlockAlign     = 2*wfm.nChannels;
    wfm.wFormatTag      = 1;

    // set up the buffer
	DSBUFFERDESC desc;

	desc.dwSize         = sizeof(desc);
	if (IsGlobal==1) {
		desc.dwFlags        = DSBCAPS_CTRLVOLUME | DSBCAPS_GLOBALFOCUS | DSBCAPS_CTRLPAN;
	} else {
		desc.dwFlags		= DSBCAPS_CTRLVOLUME | DSBCAPS_CTRLPAN;
	}
	desc.lpwfxFormat    = &wfm;
	desc.dwReserved     = 0;

	desc.dwBufferBytes  = BUFSIZE*2;

	pDS->CreateSoundBuffer(&desc, &pDSB, NULL );

    // fill the buffer

    DWORD   pos = 0;
    int     sec = 0;
    int     ret = 1;
    DWORD   size = BUFSIZE*2;

    char    *buf;

    pDSB->Lock(0, size, (LPVOID*)&buf, &size, NULL, NULL, DSBLOCK_ENTIREBUFFER);
    

	// now read in the bits
    while(ret && pos<size)
    {
        ret = ov_read(&vf, buf+pos, size-pos, 0, 2, 1, &sec);
        pos += ret;
    }

	if ( (pos==size) ) {
		shorty=false;
	} else {
		shorty=pos;
	}

	//fprintf(stream,"pos=%i,ret=%i,shorty=%i\n",pos,ret,shorty);
	//closestream();

	pDSB->Unlock( buf, size, NULL, NULL );

    nCurSection         =
    nLastSection        = 0;

    return bFileOpened = true;
}

void OggPlayer::Close()
{
    bFileOpened = false;
	nLastSection=0;

	try
	{
		if (pDSB) 
		{	
			pDSB->Release(); 
		}
	} 
	catch(...)
	{
	}
	
}


void OggPlayer::Play(bool loop)
{
    if (!bInitialized)
        return;

    if (!bFileOpened)
        return;

    // play looping because we will fill the buffer
    pDSB->Play(0,0,DSBPLAY_LOOPING);    

    bLoop = loop;
    bDone = false;
    bAlmostDone = false;
}

void OggPlayer::Stop()
{
    if (!bInitialized)
        return;

    if (!bFileOpened)
        return;

    pDSB->Stop();
}

int OggPlayer::IsPlaying()
{
	//return !bDone;
	DWORD dwStatus;
	HRESULT result=pDSB->GetStatus(&dwStatus);
	
	//openstream();
	//fprintf(stream,"bDone=%i\n",bDone);
	//closestream();

	return ( ((dwStatus & DSBSTATUS_PLAYING)==DSBSTATUS_PLAYING) &&
		(!bDone) );
}

void OggPlayer::Update()
{
	//FILE *f;
	//openstream();

	//fprintf(stream,"Beginning update...\n");
	DWORD pos;
	LONG vol;
	static int index=0;
	
	pDSB->GetCurrentPosition(&pos, NULL);

	pDSB->GetVolume(&vol);


	if (shorty) {
		//fprintf(stream, "It's a shorty, killing.\n");
		if (pos>=shorty) bDone=true;
	} else {


		nCurSection = (pos<BUFSIZE?0:1);
		
		index+=1;
		
		// section changed?
		if (nCurSection != nLastSection)
		{
			if (bDone && !bLoop) {
				Stop();
				return;
			}

			// gotta use this trick 'cause otherwise there wont be played all bits
			if (bAlmostDone && !bLoop) {
				//fprintf(stream,"Close to done.\n");
				OggPlayer::Stop();
				bDone = true;
			}

			DWORD   size = BUFSIZE;
			char    *buf;

			// fill the section we just left
			pDSB->Lock( nLastSection*BUFSIZE, size, (LPVOID*)&buf, &size, NULL, NULL, 0 );

			DWORD   pos = 0;
			int     sec = 0;
			int     ret = 1;
                
			while(ret && pos<size)
			{
				ret = ov_read(&vf, buf+pos, size-pos, 0, 2, 1, &sec);
				pos += ret;
			}

			// reached the and?
			if (!ret && bLoop)
			{
				// we are looping so restart from the beginning
				// NOTE: sound with sizes smaller than BUFSIZE may be cut off
				ret = 1;
				ov_pcm_seek(&vf, 0);
				while(ret && pos<size)
				{
					ret = ov_read(&vf, buf+pos, size-pos, 0, 2, 1, &sec);
					pos += ret;
				}
			}
			else if (!ret && !(bLoop))
			{
				// not looping so fill the rest with 0
				while(pos<size)
				{ *(buf+pos)='\0'; pos ++; }

				// and say that after the current section no other section follows
				bAlmostDone = true;
				//fprintf(stream,"Almost done!\n");
			}
			pDSB->Unlock( buf, size, NULL, NULL );
    
			nLastSection = nCurSection;
		}
	}
	//closestream();
}
