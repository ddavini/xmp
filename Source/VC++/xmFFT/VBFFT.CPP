#include "stdafx.h"
#include "xmFFT.h"
#include "windows.h"
#include "math.h"

#ifndef M_PI
#define M_PI        3.14159265358979323846
#endif

#ifndef M_PI_2
#define M_PI_2      1.57079632679489661923
#endif

float m_Top = 0;

long IntExpo(long Number,int Expo) {

	int i;
	long Store;
	Store = Number;

	for (i = 1; i < Expo; i++) {
		Number = Number * Store;
	}
	return Number;
}


//-----------------------------------
// Need the Number of Bits for a Value
// i.e. value 8 needs 3bits (2^3=8)!; 
//---------------------------------------
unsigned int NumberOfBitsNeeded(unsigned int iValue)
{
   int i;
   // iValue is Power of two always  
   for (i=1; i<=32; i++)
   {
   	if (iValue & (1<<i)) return i;
   }
   return 0;

}

bool IsPowerOfTwo(long X) {
    if (X < 2) { return false; }
    if ((X && (X - 1)) == false) { return true; }
	return false;
}

//-------------------------------------------------------
// Bit Reversing is needed in the FFT 
//-------------------------------------------------------

unsigned int ReverseBits(unsigned int iValue, unsigned int iNumberOfBits)
{
   unsigned int i,iReversedValue=0;
    
   for (i=0; i<iNumberOfBits; i++)
   {
        iReversedValue <<= 1;  // shift iValue with one to the left
   		iReversedValue = (iReversedValue ) | (iValue & 1);  //Set lsb of iReversedValue with lsb of iValue
        iValue >>= 1;  // shift iValue
   }  
   return iReversedValue;
}


float Abs(float Number) {
	
	if (Number < 0) {
		Number = Number * -1;
	}
	return Number;
}

void prepareMDCT (float *data,long Max,int Barre = 32) {
	
	
	int i = 0;

	for (i = 0; i <= Barre; i++) {
		if (data[i] > m_Top) {
			m_Top = data[i];
		}
	}

	for (i = 0; i <= Barre; i++) {
		data[i] = data[i] / (m_Top);
	}

}

/*''''''GPL FFT from XSpectrum''''''''''
inline void
swap( float &a, float &b )
{
  float t = a;  a = b;  b = t;
}

inline long double
sqr( long double arg )
{
  return arg * arg;
}

void fft( float data[], int N, int isign )
{
  int n = N << 1;
  int i, j, m;
  
  // bit reversal section 
  
  j = 0;
  for( i = 0; i < n; i += 2 ) {
    if( j > i ) {
      swap( data[j], data[i] );
      swap( data[j+1], data[i+1] );
    }
    m = N;
    while( m >= 2 && j >= m ) {
      j -= m;
      m >>= 1;
    }
    j += m;
  }
  
  // Daniel-Lanczos section 
  
  long double theta, wr, wpr, wpi, wi, wtemp;
  float tempr, tempi;
  for( int mmax = 2; n > mmax; ) {
    int istep = mmax << 1;
    theta = isign * ( 2.0*M_PI / mmax );
    wpr = -2.0 * sqr( sin( 0.5*theta ) );
    wpi = sin( theta );
    wr = 1.0;
    wi = 0.0;
    for( m = 0; m < mmax; m += 2 ) {
      for( i = m; i < n; i += istep ) {
	j = i + mmax;
	tempr = wr*data[j] - wi*data[j+1];
	tempi = wr*data[j+1] + wi*data[j];
	data[j] = data[i] - tempr;
	data[j+1] = data[i+1] - tempi;
	data[i] += tempr;
	data[i+1] += tempi;
      }
      wr = (wtemp=wr)*wpr - wi*wpi + wr;
      wi = wi*wpr + wtemp*wpi + wi;
    }
    mmax = istep;
  }
  
  // normalisation section 
  
  long double sqrtN = sqrt( N );
  for( i = 0; i < n; i++ )
    data[i] /= sqrtN;
}
''''''End - GPL FFT from XSpectrum'''''''''' 
*/

void GetResult(float *pfImagBuffer, float *pfRealBuffer,int iNum)
{
	// abs(z) = sqrt(Re(z)^2+Im(z)^2);

	int i;;
	for (i=0;i< iNum;i++)
	{
		pfRealBuffer[i] = (float)sqrt((pfRealBuffer[i]*pfRealBuffer[i])+
						    		  (pfImagBuffer[i]*pfImagBuffer[i]));
		
	}
}

void XMFFT_API Init()
{
	m_Top = 1;
}

void XMFFT_API FFT(unsigned int iNumOfSamples, float *pfReal, float *pfImag,int Barre = 32)
{
	unsigned int i,M;
	unsigned int iCol = 1; //start with first column
	unsigned int iNoS2 = iNumOfSamples / 2;
    unsigned int iKnotNumber = 0, iKnotIndex;
    unsigned int iNumberOfBits =0;
    double  dExpValue,dTmpReal, dTmpImag, dCos, dSin, dArg;

	// Formular to calculate Xl(k) = X(l-1)(k) + W^P * X(l-1)(k + N/2^l)
	// index von k mit lambda bits = Spalten NxN Matrix mit N=2^lamda
	// i Col =  Column l
	// iKnot =Knot k
	// N = Number of Samples
	// Distance between two dual Knoten N/2^l   = N<<l 
	// p shift of M = (lambda-l) to the right the do BitReversing
	       
        iNumberOfBits = NumberOfBitsNeeded(iNumOfSamples);
        if(!iNumberOfBits) return;
        

        for (iCol=1;iCol<=iNumberOfBits;iCol++)
        {
            
            do {
        	for (i=1; i<= iNoS2; i++)
        	{
        		M = (iKnotNumber>>(iNumberOfBits - iCol));
        		dExpValue = (double) ReverseBits(M,iNumberOfBits); 
        		// ---> Here is the complex calculation
        		  dArg = M_PI_2 * dExpValue / iNumOfSamples;
        		  dCos = cos (dArg);
        		  dSin = sin (dArg);
        		  dTmpReal = pfReal[iKnotNumber+iNoS2]* dCos + pfImag[iKnotNumber+iNoS2] * dSin;
        		  dTmpImag = pfImag[iKnotNumber+iNoS2]* dCos - pfReal[iKnotNumber+iNoS2] * dSin;
        		  pfReal[iKnotNumber+iNoS2] = (float)(pfReal[iKnotNumber] - dTmpReal);
        		  pfImag[iKnotNumber+iNoS2] = (float)(pfImag[iKnotNumber] - dTmpImag);
        		  pfReal[iKnotNumber] = (float)(pfReal[iKnotNumber] + dTmpReal);
        		  pfImag[iKnotNumber] = (float)(pfImag[iKnotNumber] + dTmpImag);
        		// --> End of complex calculation  
        		iKnotNumber++;
        	}
        	iKnotNumber += iNoS2;
            }while (iKnotNumber < iNumOfSamples);
            iNoS2/=2;
            iKnotNumber=0;
        }
        // now do the BitReversing
        for (iKnotNumber=0;iKnotNumber<iNumOfSamples;iKnotNumber++)
        {
        	iKnotIndex = ReverseBits(iKnotNumber, iNumberOfBits);
        	if (iKnotIndex > iKnotNumber)
        	{
        		dTmpReal = pfReal[iKnotNumber];
        		dTmpImag = pfImag[iKnotNumber];
        		pfReal[iKnotNumber] = pfReal[iKnotIndex];
        		pfImag[iKnotNumber] = pfImag[iKnotIndex];
        		pfReal[iKnotIndex] = (float)dTmpReal;
        		pfImag[iKnotIndex] = (float)dTmpImag;
       		}
        }
		GetResult(pfImag,pfReal,iNumOfSamples);
		prepareMDCT(pfReal,iNumOfSamples,Barre);
}

//obsolete 
int XMFFT_API GetxmFFTVersion()
{
	return 9999;	
}
