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
    23 , 4 , 46 , 21 , 27 , 8 , 29 , 16 , 6 , 0
      , 31 , 10 , 37 , 24 , 36 , 9 , 1 , 14 , 48 , 30
      , 43 , 11 , 19 , 40 , 3 , 45 , 47 , 33 , 25 , 18
      , 28 , 17 , 32 , 22 , 41 , 42 , 15 , 39 , 38 , 12
      , 20 , 5 , 13 , 35 , 34 , 7 , 49 , 2 , 44 , 26
  };
  integers licInt[splitSize] ={
    -393303906 , -1870215468 , -459885838 , -1700352348 , -191076732 , -2105363732 , -1565862186 , 1458356910 , -2105367916 , -1297193834
      , -257107350 , 1925897318 , -221868322 , -1461792024 , 1922358368 , -697276176 , -493581068 , 1751147110 , -455288180 , -1802343190
      , -2070351164 , -2104859450 , -597761308 , -2105376126 , -623088528 , -865154326 , 2054876766 , -2071821694 , 1819188336 , 1921290882
      , -1403612452 , -1997215074 , -1763679576 , -897132386 , -892807578 , 1688912596 , -255169814 , -2004181820 , 1855626326 , -959263510
      , -188028310 , 1616818858 , 1450242254 , -966865188 , -392141694 , -1734040356 , -228423998 , -1368759166 , -1293097852 , -2031450402
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
OrigLicKey = "AAAAB3NzaC1yc2EAAAABJQAAAIB9zFJqkkUQOTGtOuzD5mVxbNED86nmu5exK2WYjcCH0nJ9BrvYnI+Vng/c3ne9u6IJ5uezRWSMnRRLA25WOlCeue7fFmnr8N763104T1pKrq/nUY/0gx8+x48kufic+NM7oGcybyLBvYAAtioTjaU23kdeWav+oCuCyYKrA2Pt/w=="
