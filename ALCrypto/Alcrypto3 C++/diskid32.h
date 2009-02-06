//  diskid32.h

//  for displaying the details of hard drives in 

//  06/11/2000  Lynn McGuire  written with many contributions from others,
//                            IDE drives only under Windows NT/2K and 9X,
//                            maybe SCSI drives later

/**********************************************************************************************
 * Change Log
 * ==========
 *
 * Date (MM/DD/YY)  Author      Description       
 * ---------------  ----------- --------------------------------------------------------------
 * 03/13/06         J.D.M.      Ported to C++
 *
 **********************************************************************************************/

#ifndef ALCRYPTO_DISKID32_H
#define ALCRYPTO_DISKID32_H

#ifdef __cplusplus
extern "C" {
#endif

extern LRESULT WINAPI ALGetHardDriveFirmware(mUDT* pU);

#ifdef __cplusplus
}
#endif

#endif
