#include "..\\..\\minifmod160\\lib\\minifmod.h"

void memseek(unsigned int handle, int pos, signed char mode);
int memread(void *buffer, int size, unsigned int handle);
void memclose(unsigned int handle);
unsigned int memopen(char *name);
int memtell(unsigned int handle);
//X-MaD 20*01*02
signed char MOD_init();
void freeMOD(FMUSIC_MODULE* module);
FMUSIC_MODULE* mod_open(const char* pszName);
