/* diskid32.cpp */

/* for displaying the details of hard drives in */

/* 06/11/2000  Lynn McGuire  written with many contributions from others,    */
/*                           IDE drives only under Windows NT/2K and 9X,     */
/*                           maybe SCSI drives later                         */

/**********************************************************************************************
 * Change Log
 * ==========
 *
 * Date (MM/DD/YY)  Author      Description       
 * ---------------  ----------- --------------------------------------------------------------
 * 03/13/06         J.D.M.      Ported to C++
 *
 **********************************************************************************************/

/*#define PRINTING_TO_CONSOLE_ALLOWED*/ /* Define this one to let PrintIdeInfo() print to the console */

#include <stdlib.h>
#include <stdio.h>
#include <windows.h>

#include "alcrypto3.h"
#include "diskid32.h"
#include "winio.h"
#include "port32.h"
#include "bignum.h"

/* Required to ensure correct PhysicalDrive IOCTL structure setup */
#pragma pack(1)

/* Max number of drives assuming primary/secondary, master/slave topology */
#define  MAX_IDE_DRIVES  4
#define  IDENTIFY_BUFFER_SIZE  512

/* IOCTL commands */
#define  DFP_GET_VERSION          0x00074080
#define  DFP_SEND_DRIVE_COMMAND   0x0007c084
#define  DFP_RECEIVE_DRIVE_DATA   0x0007c088

#define  FILE_DEVICE_SCSI              0x0000001b
#define  IOCTL_SCSI_MINIPORT_IDENTIFY  ((FILE_DEVICE_SCSI << 16) + 0x0501)
#define  IOCTL_SCSI_MINIPORT 0x0004D008  /* see NTDDSCSI.H for definition */

/* GETVERSIONOUTPARAMS contains the data returned from the */
/* Get Driver Version function.                            */
typedef struct _GETVERSIONOUTPARAMS {
  BYTE bVersion;       /* Binary driver version. */
  BYTE bRevision;      /* Binary driver revision. */
  BYTE bReserved;      /* Not used. */
  BYTE bIDEDeviceMap;  /* Bit map of IDE devices. */
  DWORD fCapabilities; /* Bit mask of driver capabilities. */
  DWORD dwReserved[4]; /* For future use. */
} GETVERSIONOUTPARAMS, *PGETVERSIONOUTPARAMS, *LPGETVERSIONOUTPARAMS;

/* Bits returned in the fCapabilities member of GETVERSIONOUTPARAMS */
#define  CAP_IDE_ID_FUNCTION             1  /* ATA ID command supported */
#define  CAP_IDE_ATAPI_ID                2  /* ATAPI ID command supported */
#define  CAP_IDE_EXECUTE_SMART_FUNCTION  4  /* SMART commannds supported */

/* IDE registers */
typedef struct _IDEREGS {
  BYTE bFeaturesReg;       /* Used for specifying SMART "commands". */
  BYTE bSectorCountReg;    /* IDE sector count register */
  BYTE bSectorNumberReg;   /* IDE sector number register */
  BYTE bCylLowReg;         /* IDE low order cylinder value */
  BYTE bCylHighReg;        /* IDE high order cylinder value */
  BYTE bDriveHeadReg;      /* IDE drive/head register */
  BYTE bCommandReg;        /* Actual IDE command. */
  BYTE bReserved;          /* Reserved for future use. Must be zero. */
} IDEREGS, *PIDEREGS, *LPIDEREGS;

/* SENDCMDINPARAMS contains the input parameters for the */
/* Send Command to Drive function.                       */
typedef struct _SENDCMDINPARAMS {
  DWORD     cBufferSize;   /* Buffer size in bytes */
  IDEREGS   irDriveRegs;   /* Structure with drive register values. */
  BYTE bDriveNumber;       /* Physical drive number to send  */
  /* command to (0,1,2,3). */
  BYTE bReserved[3];       /* Reserved for future expansion. */
  DWORD     dwReserved[4]; /* For future use. */
  BYTE      bBuffer[1];    /* Input buffer. */
} SENDCMDINPARAMS, *PSENDCMDINPARAMS, *LPSENDCMDINPARAMS;

/* Valid values for the bCommandReg member of IDEREGS. */
#define  IDE_ATAPI_IDENTIFY  0xA1  /* Returns ID sector for ATAPI. */
#define  IDE_ATA_IDENTIFY    0xEC  /* Returns ID sector for ATA. */

/* Status returned from driver */
typedef struct _DRIVERSTATUS {
  BYTE  bDriverError;   /* Error code from driver, or 0 if no error. */
  BYTE  bIDEStatus;     /* Contents of IDE Error register. */
  /*  Only valid when bDriverError is SMART_IDE_ERROR. */
  BYTE  bReserved[2];   /* Reserved for future expansion. */
  DWORD  dwReserved[2]; /* Reserved for future expansion. */
} DRIVERSTATUS, *PDRIVERSTATUS, *LPDRIVERSTATUS;

/* Structure returned by PhysicalDrive IOCTL for several commands */
typedef struct _SENDCMDOUTPARAMS {
  DWORD         cBufferSize;  /* Size of bBuffer in bytes */
  DRIVERSTATUS  DriverStatus; /* Driver status structure. */
  BYTE          bBuffer[1];   /* Buffer of arbitrary length in which to store the data read from the drive. */
} SENDCMDOUTPARAMS, *PSENDCMDOUTPARAMS, *LPSENDCMDOUTPARAMS;

#define  SENDIDLENGTH  sizeof (SENDCMDOUTPARAMS) + IDENTIFY_BUFFER_SIZE

/* The following struct defines the interesting part of the IDENTIFY buffer: */
typedef struct _IDSECTOR {
  USHORT  wGenConfig;
  USHORT  wNumCyls;
  USHORT  wReserved;
  USHORT  wNumHeads;
  USHORT  wBytesPerTrack;
  USHORT  wBytesPerSector;
  USHORT  wSectorsPerTrack;
  USHORT  wVendorUnique[3];
  CHAR    sSerialNumber[20];
  USHORT  wBufferType;
  USHORT  wBufferSize;
  USHORT  wECCSize;
  CHAR    sFirmwareRev[8];
  CHAR    sModelNumber[40];
  USHORT  wMoreVendorUnique;
  USHORT  wDoubleWordIO;
  USHORT  wCapabilities;
  USHORT  wReserved1;
  USHORT  wPIOTiming;
  USHORT  wDMATiming;
  USHORT  wBS;
  USHORT  wNumCurrentCyls;
  USHORT  wNumCurrentHeads;
  USHORT  wNumCurrentSectorsPerTrack;
  ULONG   ulCurrentSectorCapacity;
  USHORT  wMultSectorStuff;
  ULONG   ulTotalAddressableSectors;
  USHORT  wSingleWordDMA;
  USHORT  wMultiWordDMA;
  BYTE    bReserved[128];
} IDSECTOR, *PIDSECTOR;

typedef struct _SRB_IO_CONTROL {
  ULONG HeaderLength;
  UCHAR Signature[8];
  ULONG Timeout;
  ULONG ControlCode;
  ULONG ReturnCode;
  ULONG Length;
} SRB_IO_CONTROL, *PSRB_IO_CONTROL;

static char HardDriveSerialNumber [1024];

/* Define global buffers. */
BYTE IdOutCmd [sizeof (SENDCMDOUTPARAMS) + IDENTIFY_BUFFER_SIZE - 1];

static int  ReadPhysicalDriveInNT(void);
static BOOL DoIDENTIFY (HANDLE, PSENDCMDINPARAMS, PSENDCMDOUTPARAMS, BYTE, BYTE, PDWORD);
static int  ReadDrivePortsInWin9X (void);
static int  ReadIdeDriveAsScsiDriveInNT (void);
static void PrintIdeInfo (int drive, DWORD diskdata [256]);
static char *ConvertToString (DWORD diskdata [256], int firstIndex, int lastIndex);

/*---------------------------------------------------------------------------*/
/* ALGetHardDriveFirmware                                                    */
/*---------------------------------------------------------------------------*/
LRESULT WINAPI ALGetHardDriveFirmware(mUDT* pU)
{
  int  done = FALSE;
  char *p = HardDriveSerialNumber;

  strcpy (HardDriveSerialNumber, "");
  
  /* this works under WinNT4 or Win2K if you have admin rights */
  done = ReadPhysicalDriveInNT ();
  
  if ( ! done) {
    /* this should work in WinNT or Win2K if previous did not work   */
    /* this is kind of a backdoor via the SCSI mini port driver into */
    /* the IDE drives                                                */
    done = ReadIdeDriveAsScsiDriveInNT ();
  }
  
  if ( ! done) {
    /* this works under Win9X and calls WINIO.DLL */
    done = ReadDrivePortsInWin9X ();
  }

  /* ignore first 5 characters from western digital hard drives if */
  /* the first four characters are WD-W                            */
  /* if ( ! strncmp (HardDriveSerialNumber, "WD-W", 4)) {
    p += 5;
  } */
    
  strncpy(pU->cArray, p, 30);
  pU->lformat = wf1e;

  return 30;
} /* ALGetHardDriveFirmware */

/*---------------------------------------------------------------------------*/
/* ReadPhysicalDriveInNT                                                     */
/*---------------------------------------------------------------------------*/
static int ReadPhysicalDriveInNT(void)
{
  int done = FALSE;
  int drive = 0;

  for (drive = 0; drive < MAX_IDE_DRIVES; drive++) {
    HANDLE hPhysicalDriveIOCTL = 0;

    /* Try to get a handle to PhysicalDrive IOCTL, report failure */
    /* and exit if can't.                                         */
    char driveName [256];

    sprintf (driveName, "\\\\.\\PhysicalDrive%d", drive);

    /* Windows NT, Windows 2000, must have admin rights */
    hPhysicalDriveIOCTL = CreateFile (driveName, GENERIC_READ | GENERIC_WRITE, 
                                      FILE_SHARE_READ | FILE_SHARE_WRITE, NULL,
                                      OPEN_EXISTING, 0, NULL);
    /* if (hPhysicalDriveIOCTL == INVALID_HANDLE_VALUE) {                                           */
    /*    printf ("Unable to open physical drive %d, error code: 0x%lX\n", drive, GetLastError ()); */
    /* }                                                                                            */

    if (hPhysicalDriveIOCTL != INVALID_HANDLE_VALUE) {
      GETVERSIONOUTPARAMS VersionParams;
      DWORD               cbBytesReturned = 0;

      /* Get the version, etc of PhysicalDrive IOCTL */
      memset ((void*) &VersionParams, 0, sizeof(VersionParams));

      if ( ! DeviceIoControl(hPhysicalDriveIOCTL, DFP_GET_VERSION, NULL, 0,
                             &VersionParams, sizeof(VersionParams), &cbBytesReturned, NULL) ) {         
        /* printf ("DFP_GET_VERSION failed for drive %d\n", i); */
        /* continue;                                            */
      }

      /* If there is a IDE device at number "i" issue commands */
      /* to the device                                         */
      if (VersionParams.bIDEDeviceMap > 0) {
        BYTE             bIDCmd = 0;   /* IDE or ATAPI IDENTIFY cmd */
        SENDCMDINPARAMS  scip;
        /* SENDCMDOUTPARAMS OutCmd; */

        /* Now, get the ID sector for all IDE devices in the system.  */
        /* If the device is ATAPI use the IDE_ATAPI_IDENTIFY command, */
        /* otherwise use the IDE_ATA_IDENTIFY command                 */
        bIDCmd = (VersionParams.bIDEDeviceMap >> drive & 0x10) ? IDE_ATAPI_IDENTIFY : IDE_ATA_IDENTIFY;

        memset (&scip, 0, sizeof(scip));
        memset (IdOutCmd, 0, sizeof(IdOutCmd));

        if ( DoIDENTIFY (hPhysicalDriveIOCTL, &scip, (PSENDCMDOUTPARAMS)&IdOutCmd, 
                         (BYTE)bIDCmd, (BYTE)drive, &cbBytesReturned)) {
          DWORD diskdata [256];
          int ijk = 0;
          USHORT *pIdSector = (USHORT *)((PSENDCMDOUTPARAMS) IdOutCmd) -> bBuffer;

          for (ijk = 0; ijk < 256; ijk++) {
            diskdata [ijk] = pIdSector [ijk];
          }

          /* copy the hard driver serial number to the buffer */
          strcpy (HardDriveSerialNumber, ConvertToString (diskdata, 10, 19));
          PrintIdeInfo (drive, diskdata);

          done = TRUE;
        }
      }

      CloseHandle (hPhysicalDriveIOCTL);
    }
  }

  return done;
} /* ReadPhysicalDriveInNT */

/*---------------------------------------------------------------------------*/
/* DoIDENTIFY                                                                */
/*---------------------------------------------------------------------------*/
/* FUNCTION: Send an IDENTIFY command to the drive                           */
/* bDriveNum = 0-3                                                           */
/* bIDCmd = IDE_ATA_IDENTIFY or IDE_ATAPI_IDENTIFY                           */
/*---------------------------------------------------------------------------*/
static BOOL DoIDENTIFY (HANDLE hPhysicalDriveIOCTL, PSENDCMDINPARAMS pSCIP,
                        PSENDCMDOUTPARAMS pSCOP, BYTE bIDCmd, BYTE bDriveNum,
                        PDWORD lpcbBytesReturned)
{
  /* Set up data structures for IDENTIFY command. */
  pSCIP -> cBufferSize = IDENTIFY_BUFFER_SIZE;
  pSCIP -> irDriveRegs.bFeaturesReg = 0;
  pSCIP -> irDriveRegs.bSectorCountReg = 1;
  pSCIP -> irDriveRegs.bSectorNumberReg = 1;
  pSCIP -> irDriveRegs.bCylLowReg = 0;
  pSCIP -> irDriveRegs.bCylHighReg = 0;

  /* Compute the drive number. */
  pSCIP -> irDriveRegs.bDriveHeadReg = 0xA0 | ((bDriveNum & 1) << 4);

  /* The command can either be IDE identify or ATAPI identify. */
  pSCIP -> irDriveRegs.bCommandReg = bIDCmd;
  pSCIP -> bDriveNumber = bDriveNum;
  pSCIP -> cBufferSize = IDENTIFY_BUFFER_SIZE;

  return ( DeviceIoControl (hPhysicalDriveIOCTL, DFP_RECEIVE_DRIVE_DATA,
           (LPVOID) pSCIP, sizeof(SENDCMDINPARAMS) - 1, (LPVOID) pSCOP,
           sizeof(SENDCMDOUTPARAMS) + IDENTIFY_BUFFER_SIZE - 1,
           lpcbBytesReturned, NULL) );
} /* DoIDENTIFY */

/*---------------------------------------------------------------------------*/
/* ReadDrivePortsInWin9X                                                     */
/*---------------------------------------------------------------------------*/
static int ReadDrivePortsInWin9X (void)
{
  int done = FALSE;
  int drive = 0;

  InitializeWinIo ();   

  /* Get IDE Drive info from the hardware ports */
  /* loop thru all possible drives              */
  for (drive = 0; drive < 8; drive++) {
    DWORD diskdata [256];
    WORD  baseAddress = 0;   //  Base address of drive controller 
    DWORD portValue = 0;
    int waitLoop = 0;
    int index = 0;

    switch (drive / 2) {
      case 0:
        baseAddress = 0x1f0;
        break;
      case 1:
        baseAddress = 0x170;
        break;
      case 2:
        baseAddress = 0x1e8;
        break;
      case 3:
        baseAddress = 0x168;
        break;
      default:
        /* Should not happen */
        break;
    }

    /* Wait for controller not busy */
    waitLoop = 100000;
    while (--waitLoop > 0) {
      GetPortVal ((WORD) (baseAddress + 7), &portValue, (BYTE) 1);
      /* drive is ready */
      if ((portValue & 0x40) == 0x40) {
        break;
      }
      /* previous drive command ended in error */
      if ((portValue & 0x01) == 0x01) {
        break;
      }
    }

    if (waitLoop < 1) {
      continue;
    }

    /* Set Master or Slave drive */
    if ((drive % 2) == 0) {
      SetPortVal ((WORD) (baseAddress + 6), 0xA0, 1);
    }
    else {
      SetPortVal ((WORD) (baseAddress + 6), 0xB0, 1);
    }

    /* Get drive info data */
    SetPortVal ((WORD) (baseAddress + 7), 0xEC, 1);

    /* Wait for data ready */
    waitLoop = 100000;
    while (--waitLoop > 0) {
      GetPortVal ((WORD) (baseAddress + 7), &portValue, 1);
      /* see if the drive is ready and has it's info ready for us */
      if ((portValue & 0x48) == 0x48) {
        break;
      }
      /* see if there is a drive error */
      if ((portValue & 0x01) == 0x01) {
        break;
      }
    }

    /* check for time out or other error */
    if (waitLoop < 1 || portValue & 0x01) {
      continue;
    }

    /* read drive id information */
    for (index = 0; index < 256; index++) {
      diskdata [index] = 0; /* init the space */
      GetPortVal (baseAddress, &(diskdata [index]), 2);
    }

    /* copy the hard driver serial number to the buffer */
    strcpy (HardDriveSerialNumber, ConvertToString (diskdata, 10, 19));
    PrintIdeInfo (drive, diskdata);

    done = TRUE;
  }
  
  ShutdownWinIo ();

  return done;
} /* ReadDrivePortsInWin9X */

/*---------------------------------------------------------------------------*/
/* ReadIdeDriveAsScsiDriveInNT                                               */
/*---------------------------------------------------------------------------*/
static int ReadIdeDriveAsScsiDriveInNT (void)
{
  int done = FALSE;
  int controller = 0;

  for (controller = 0; controller < 2; controller++) {
    HANDLE hScsiDriveIOCTL = 0;
    char   driveName [256];

    /* Try to get a handle to PhysicalDrive IOCTL, report failure */
    /* and exit if can't.                                         */
    sprintf (driveName, "\\\\.\\Scsi%d:", controller);

    /* Windows NT, Windows 2000, any rights should do */
    hScsiDriveIOCTL = CreateFile (driveName, GENERIC_READ | GENERIC_WRITE,
                                  FILE_SHARE_READ | FILE_SHARE_WRITE, NULL,
                                  OPEN_EXISTING, 0, NULL);
    /* if (hScsiDriveIOCTL == INVALID_HANDLE_VALUE) {                      */
    /*   printf ("Unable to open SCSI controller %d, error code: 0x%lX\n", */
    /*            controller, GetLastError ());                            */
    /* }                                                                   */

    if (hScsiDriveIOCTL != INVALID_HANDLE_VALUE) {
      int drive = 0;

      for (drive = 0; drive < 2; drive++) {
        char            buffer [sizeof (SRB_IO_CONTROL) + SENDIDLENGTH];
        SRB_IO_CONTROL  *p = (SRB_IO_CONTROL *) buffer;
        SENDCMDINPARAMS *pin = (SENDCMDINPARAMS *) (buffer + sizeof (SRB_IO_CONTROL));
        DWORD           dummy;

        memset (buffer, 0, sizeof (buffer));
        p -> HeaderLength = sizeof (SRB_IO_CONTROL);
        p -> Timeout = 10000;
        p -> Length = SENDIDLENGTH;
        p -> ControlCode = IOCTL_SCSI_MINIPORT_IDENTIFY;
        strncpy ((char *) p -> Signature, "SCSIDISK", 8);
        
        pin -> irDriveRegs.bCommandReg = IDE_ATA_IDENTIFY;
        pin -> bDriveNumber = drive;

        if (DeviceIoControl (hScsiDriveIOCTL, IOCTL_SCSI_MINIPORT, buffer,
                             sizeof (SRB_IO_CONTROL) + sizeof (SENDCMDINPARAMS) - 1,
                             buffer, sizeof (SRB_IO_CONTROL) + SENDIDLENGTH,
                             &dummy, NULL)) {
          SENDCMDOUTPARAMS *pOut = (SENDCMDOUTPARAMS *) (buffer + sizeof (SRB_IO_CONTROL));
          IDSECTOR *pId = (IDSECTOR *) (pOut -> bBuffer);

          if (pId -> sModelNumber [0]) {
            DWORD diskdata [256];
            int ijk = 0;
            USHORT *pIdSector = (USHORT *) pId;
            
            for (ijk = 0; ijk < 256; ijk++) {
              diskdata [ijk] = pIdSector [ijk];
            }

            /* copy the hard driver serial number to the buffer */
            strcpy (HardDriveSerialNumber, ConvertToString (diskdata, 10, 19));
            PrintIdeInfo (controller * 2 + drive, diskdata);

            done = TRUE;
          }
        }
      }
      CloseHandle (hScsiDriveIOCTL);
    }
  }

  return done;
} /* ReadIdeDriveAsScsiDriveInNT */

/*---------------------------------------------------------------------------*/
/* PrintIdeInfo                                                              */
/*---------------------------------------------------------------------------*/
static void PrintIdeInfo (int drive, DWORD diskdata [256])
{
#ifdef PRINTING_TO_CONSOLE_ALLOWED
  switch (drive / 2) {
    case 0:
      printf ("\nPrimary Controller - ");
      break;
    case 1:
      printf ("\nSecondary Controller - ");
      break;
    case 2:
      printf ("\nTertiary Controller - ");
      break;
    case 3:
      printf ("\nQuaternary Controller - ");
      break;
    default:
      /* Should not happen */
      break;
  }

  switch (drive % 2) {
    case 0: printf ("Master drive\n\n");
      break;
    case 1: printf ("Slave drive\n\n");
      break;
    default:
      /* Should not happen */
      break;
  }

  printf ("Drive Model Number________________: %s\n",
  ConvertToString (diskdata, 27, 46));
  printf ("Drive Serial Number_______________: %s\n",
  ConvertToString (diskdata, 10, 19));
  printf ("Drive Controller Revision Number__: %s\n",
  ConvertToString (diskdata, 23, 26));

  printf ("Controller Buffer Size on Drive___: %u bytes\n",
  diskdata [21] * 512);

  printf ("Drive Type________________________: ");
  if (diskdata [0] & 0x0080) {
    printf ("Removable\n");
  }
  else if (diskdata [0] & 0x0040) {
    printf ("Fixed\n");
  }
  else {
    printf ("Unknown\n");
  }

  printf ("Physical Geometry:     %u Cylinders     %u Heads     %u Sectors per track\n",
          diskdata [1], diskdata [3], diskdata [6]);

#endif
} /* PrintIdeInfo */

/*---------------------------------------------------------------------------*/
/* ConvertToString                                                           */
/*---------------------------------------------------------------------------*/
static char *ConvertToString (DWORD diskdata [256], int firstIndex, int lastIndex)
{
  static char string [1024];
  int         index = 0;
  int         position = 0;

  /* each integer has two characters stored in it backwards */
  for (index = firstIndex; index <= lastIndex; index++) {
    /* get high byte for 1st character */
    string [position] = (char) (diskdata [index] / 256);
    position++;

    /* get low byte for 2nd character */
    string [position] = (char) (diskdata [index] % 256);
    position++;
  }

  /* end the string */
  string [position] = '\0';

  /* cut off the trailing blanks */
  for (index = position - 1; index > 0 && ' ' == string [index]; index--) {
    string [index] = '\0';
  }

  return string;
} /* ConvertToString */
