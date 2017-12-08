//*****************************************************************************************/
// iFlySTTSApi.h
//
// author: zzh
// date :2001/2/27
// last update :2001/3/26
// company: iFLYTEK
// remarks:This head file defines the interfaces of the simple TTS service.
// ****************************************************************************************/

#ifndef _IFLY_STTSAPI_H_
#define _IFLY_STTSAPI_H_

#include "iFly_TTS.h"

#ifdef WIN32
#define STTSWINAPI WINAPI
#else 
#define STTSWINAPI
#endif

////////////////////////////////////////////////////////////////////////////////////////////
//appendix functions
//------------------------------------------------------------------------------------------
// Function name: STTSLoadLibrary
// Return value : TRUE if the function is successful; otherwise FALSE.
// Remarks      : Load STTSapi.dll and locate the  address of the functions.
//------------------------------------------------------------------------------------------
//以下是函数声明
extern BOOL STTSLoadLibrary();

extern BOOL  STTSUnloadLibrary();
 
extern BOOL (STTSWINAPI *STTSInit)();

extern BOOL (STTSWINAPI *STTSDeinit)();

extern HTTSINSTANCE (STTSWINAPI *STTSConnect)(TTSCHAR* pszSerialNumber,TTSCHAR* pszTTSServerIP);

extern BOOL (STTSWINAPI *STTSDisconnect)(HTTSINSTANCE hTTSInstance);

extern BOOL (STTSWINAPI *STTSSetParam)(HTTSINSTANCE hTTSInstance,int nType,int nParam);

extern BOOL (STTSWINAPI *STTSGetParam)(HTTSINSTANCE hTTSInstance,int nType, int* nParam);

extern BOOL (STTSWINAPI *STTSString2AudioFile)(HTTSINSTANCE hTTSInstance, TTSCHAR* pszString, TTSCHAR* pszAudioFile);

extern BOOL (STTSWINAPI *STTSFile2AudioFile)(HTTSINSTANCE hTTSInstance,TTSCHAR* pszTextFile, TTSCHAR* pszAudioFile);

extern BOOL (STTSWINAPI *STTSPlayString)(HTTSINSTANCE hTTSInstance , TTSCHAR* pszString,BOOL bASynch);

extern BOOL (STTSWINAPI *STTSPlayTextFile)(HTTSINSTANCE hTTSInstance ,TTSCHAR* pszTextFile,BOOL bASynch);

extern BOOL (STTSWINAPI *STTSPlayStop)(void );

extern BOOL (STTSWINAPI *STTSQueryPlayStatus)(int* nStatus);

extern BOOL (STTSWINAPI *STTSAbout)(int nAboutType,TTSCHAR* pszAboutInfo, int* pnInfoSize);

#endif