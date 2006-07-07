// code here is all in one lump. Seperate and Hide ...
// Advice. If you are new to ActiveLock, then get it working by just using the long license key first
// When that is working implement this as its own stage
// Remember if you produce a new key, eg a new version YOU MUST redo this step
Typical includes should be, these are the ones in the program StringHider. Not all are nescessary !
#include <iostream>      
#include <functional>    
#include <string>        
#include <fstream>       
#include <vector>        
#include <algorithm>     
#include <strstream>     
#include <iomanip>       
using namespace std;     
string GetIt(){
	typedef unsigned int integers;
	const int charsPerInt = 4;
	const int licSize     = 200;
	const int splitSize   = licSize/charsPerInt;
	const int ileft       = 0;
	int untwist[splitSize] ={
		26 , 16 , 18 , 34 , 4 , 27 , 0 , 9 , 35 , 23
		, 1 , 41 , 21 , 3 , 47 , 19 , 12 , 40 , 5 , 20
		, 29 , 45 , 39 , 13 , 30 , 37 , 8 , 14 , 10 , 42
		, 15 , 32 , 7 , 22 , 33 , 11 , 49 , 38 , 6 , 24
		, 2 , 43 , 17 , 31 , 36 , 25 , 46 , 48 , 28 , 44
	};
	integers licInt[splitSize] ={
		1720095362 , -729511318 , -1297946482 , -528158122 , -2071821694 , 1821831818 , -1499423588 , 1653251294 , -697382182 , -1503622544
		, -1872181046 , -1565481842 , -1771256696 , -2070353258 , -1695391132 , -1869944632 , -191076732 , -1738381104 , -228423998 , -1033213326
		, -1571367326 , -561450806 , -656101784 , -1569690458 , -1394711850 , 1622333546 , -2105376126 , -2105367916 , -1835868048 , 1584163012
		, -957034354 , -959540574 , -1729568546 , -427399960 , -2104859450 , -1298074426 , -186343712 , -387288442 , 1921817290 , -2064340384
		, -420974448 , -1970096918 , -1533219644 , -729117496 , 2054857380 , -1369401706 , -925865756 , 1886293678 , -695800158 , -1335833386
	};
	string stLic;
	stLic.reserve(licSize);
	for(int* cit = &untwist[0]; cit != &untwist[splitSize]; cit++)
	{
		char cstr[charsPerInt+1];
		int index = *cit;
		unsigned int i = licInt[index];
		i = _rotr( i, 1 );
		const unsigned int * pi = &i;
		const char * pc = reinterpret_cast<const char*>(pi);
		for(int ic=0; ic < charsPerInt; ic++){
			cstr[ic] = *pc++;
		}
		cstr[charsPerInt] = 0;
		stLic += cstr;
		const char* pTest = stLic.c_str();  // test only
	}
	if(ileft){ stLic = stLic.substr(0, (int)stLic.length() - (charsPerInt-ileft)); }
	return stLic;
}
// add following one of the following lines to access procedure
const char* pLic = GetIt().c_str();
string licKey = GetIt();
// The idea is not to have the following 2 lines in the program                          
// so remove them. Only here too check that the procedure works. When satisfied remove them or convert to comment
string OrigLicKey;
OrigLicKey = "AAAAB3NzaC1yc2EAAAABJQAAAIC38E0SczPYS68Q5EBjuZIEexDo+yBpWG78955aDb6KH8tsEyK61k+Qb46/KK0W0UzBKuLBGhzcC9utme7k21yMed4HbvNRdtEHoptLoNE14Wrlt2CsGRXQkf0XeMF9NNPSk1oVGtQYdFEjh41LQIgcpOrz5lY0r4hdQwCk8pIIRQ=="
