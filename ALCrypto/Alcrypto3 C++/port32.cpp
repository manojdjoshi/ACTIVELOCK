// ------------------------------------------------ //
//                  Port32 v3.0                     //
//    Direct Port Access Under Windows 9x/NT/2000   //
//        Copyright 1998-2000 Yariv Kaplan          //
//            http://www.internals.com              //
// ------------------------------------------------ //

#include <windows.h>
#include <winioctl.h>

#include "alcrypto3.h"
#include "port32.h"
#include "winio.h"
/*#include "general.h"*/

static bool CallRing0(PVOID pvRing0FuncAddr, WORD wPortAddr, PDWORD pdwPortVal, BYTE bSize);

/* These are our ring 0 functions responsible for tinkering with the hardware ports. */
/* They have a similar privilege to a Windows VxD and are therefore free to access   */
/* protected system resources (such as the page tables) and even place calls to      */
/* exported VxD services.                                                            */

/*---------------------------------------------------------------------------*/
/* Ring0GetPortVal                                                           */
/*---------------------------------------------------------------------------*/
static __declspec(naked) void Ring0GetPortVal()
{
  _asm
  {
    Cmp CL, 1
    Je ByteVal
    Cmp CL, 2
    Je WordVal
    Cmp CL, 4
    Je DWordVal

ByteVal:
    In AL, DX
    Mov [EBX], AL
    Retf

WordVal:
    In AX, DX
    Mov [EBX], AX
    Retf

DWordVal:
    In EAX, DX
    Mov [EBX], EAX
    Retf
  }
} /* Ring0GetPortVal */

/*---------------------------------------------------------------------------*/
/* Ring0SetPortVal                                                           */
/*---------------------------------------------------------------------------*/
static __declspec(naked) void Ring0SetPortVal()
{
  _asm
  {
    Cmp CL, 1
    Je ByteVal
    Cmp CL, 2
    Je WordVal
    Cmp CL, 4
    Je DWordVal

ByteVal:
    Mov AL, [EBX]
    Out DX, AL
    Retf

WordVal:
    Mov AX, [EBX]
    Out DX, AX
    Retf

DWordVal:
    Mov EAX, [EBX]
    Out DX, EAX
    Retf
  }
} /* Ring0SetPortVal */

/*---------------------------------------------------------------------------*/
/* CallRing0                                                                 */
/*---------------------------------------------------------------------------*/
/* This function makes it possible to call ring 0 code from a ring 3         */
/* application.                                                              */
/*---------------------------------------------------------------------------*/
static bool CallRing0(PVOID pvRing0FuncAddr, WORD wPortAddr, PDWORD pdwPortVal, BYTE bSize)
{
  struct GDT_DESCRIPTOR *pGDTDescriptor;
  struct GDTR gdtr;
  WORD   CallgateAddr[3];
  WORD   wGDTIndex = 1;

  _asm Sgdt [gdtr]

   /* Skip the null descriptor */
   pGDTDescriptor = (struct GDT_DESCRIPTOR *)(gdtr.dwGDTBase + 8);

  /* Search for a free GDT descriptor */
  for (wGDTIndex = 1; wGDTIndex < (gdtr.wGDTLimit / 8); wGDTIndex++) {
    if (   pGDTDescriptor->Type == 0
        && pGDTDescriptor->System == 0
        && pGDTDescriptor->DPL == 0
        && pGDTDescriptor->Present == 0) {
      /* Found one !                                               */
      /* Now we need to transform this descriptor into a callgate. */
      /* Note that we're using selector 0x28 since it corresponds  */
      /* to a ring 0 segment which spans the entire linear address */
      /* space of the processor (0-4GB).                           */

      struct CALLGATE_DESCRIPTOR *pCallgate;

      pCallgate = (struct CALLGATE_DESCRIPTOR *) pGDTDescriptor;
      pCallgate->Offset_0_15 = LOWORD(pvRing0FuncAddr);
      pCallgate->Selector = 0x28;
      pCallgate->ParamCount = 0;
      pCallgate->Unused = 0;
      pCallgate->Type = 0xc;
      pCallgate->System = 0;
      pCallgate->DPL = 3;
      pCallgate->Present = 1;
      pCallgate->Offset_16_31 = HIWORD(pvRing0FuncAddr);

      /* Prepare the far call parameters */
      CallgateAddr[0] = 0x0;
      CallgateAddr[1] = 0x0;
      CallgateAddr[2] = (wGDTIndex << 3) | 3;

      /* Please fasten your seat belts!                     */
      /* We're about to make a hyperspace jump into RING 0. */
      _asm Mov DX, [wPortAddr]
      _asm Mov EBX, [pdwPortVal]
      _asm Mov CL, [bSize]
      _asm Call FWORD PTR [CallgateAddr]

      /* We have made it!            */
      /* Now free the GDT descriptor */
      memset(pGDTDescriptor, 0, 8);

      /* Our journey was successful. Seeya. */
      return true;
    }

    /* Advance to the next GDT descriptor */
    pGDTDescriptor++; 
  }

  /* Whoops, the GDT is full */
  return false;
} /* CallRing0 */

/*---------------------------------------------------------------------------*/
/* GetPortVal                                                                */
/*---------------------------------------------------------------------------*/
bool GetPortVal(WORD wPortAddr, PDWORD pdwPortVal, BYTE bSize)
{
  bool   Result;
  DWORD  dwBytesReturned;
  struct tagPort32Struct Port32Struct;

  if (IsNT) {
    if (!IsWinIoInitialized) {
      return false;
    }

    Port32Struct.wPortAddr = wPortAddr;
    Port32Struct.bSize = bSize;

    if (!DeviceIoControl(hDriver, IOCTL_WINIO_READPORT, &Port32Struct,
                         sizeof(struct tagPort32Struct), &Port32Struct, 
                         sizeof(struct tagPort32Struct),
                         &dwBytesReturned, NULL)) {
      return false;
    }
    else {
      *pdwPortVal = Port32Struct.dwPortVal;
    }
  }
  else {
    Result = CallRing0((PVOID)Ring0GetPortVal, wPortAddr, pdwPortVal, bSize);

    if (Result == false) {
      return false;
    }
  }

  return true;
} /* GetPortVal */

/*---------------------------------------------------------------------------*/
/* SetPortVal                                                                */
/*---------------------------------------------------------------------------*/
bool SetPortVal(WORD wPortAddr, DWORD dwPortVal, BYTE bSize)
{
  DWORD  dwBytesReturned;
  struct tagPort32Struct Port32Struct;

  if (IsNT) {
    if (!IsWinIoInitialized) {
      return false;
    }

    Port32Struct.wPortAddr = wPortAddr;
    Port32Struct.dwPortVal = dwPortVal;
    Port32Struct.bSize = bSize;

    if (!DeviceIoControl(hDriver, IOCTL_WINIO_WRITEPORT, &Port32Struct,
                         sizeof(struct tagPort32Struct), NULL, 0,
                         &dwBytesReturned, NULL)) {
      return false;
    }
  }
  else {
    return CallRing0((PVOID)Ring0SetPortVal, wPortAddr, &dwPortVal, bSize);
  }

  return true;
} /* SetPortVal */
