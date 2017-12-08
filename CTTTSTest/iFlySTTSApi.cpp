// ******************************************************************************************/
// iFlySTTSApi.cpp
// Last Update Date : 2001.3.27
// Author : zzh
// iFlyTek (c) 2001
// ******************************************************************************************/
#include "stdafx.h"
#include "iflysttsapi.h"

extern BOOL bSTTSApiLoaded ;
extern HMODULE hSTTSApiMod;

HMODULE hSTTSApiMod;
BOOL bSTTSApiLoaded = FALSE;

//////////////////////////////////////////////////////////////////////////////////////////////
////Declare points of the funtions

BOOL (STTSWINAPI *STTSInit)();
BOOL (STTSWINAPI *STTSDeinit)();
HTTSINSTANCE (STTSWINAPI *STTSConnect)(TTSCHAR* pszSerialNumber,TTSCHAR* pszTTSServerIP);
BOOL (STTSWINAPI *STTSDisconnect)(HTTSINSTANCE hTTSInstance);
BOOL (STTSWINAPI *STTSSetParam)(HTTSINSTANCE hTTSInstance,int nType,int nParam);
BOOL (STTSWINAPI *STTSGetParam)(HTTSINSTANCE hTTSInstance,int nType, int* nParam);
BOOL (STTSWINAPI *STTSString2AudioFile)(HTTSINSTANCE hTTSInstance, TTSCHAR* pszString, TTSCHAR* pszAudioFile);
BOOL (STTSWINAPI *STTSFile2AudioFile)(HTTSINSTANCE hTTSInstance,TTSCHAR* pszTextFile, TTSCHAR* pszAudioFile);
BOOL (STTSWINAPI *STTSPlayString)(HTTSINSTANCE hTTSInstance , TTSCHAR* pszString,BOOL bASynch);
BOOL (STTSWINAPI *STTSPlayTextFile)(HTTSINSTANCE hTTSInstance ,TTSCHAR* pszTextFile,BOOL bASynch);
BOOL (STTSWINAPI *STTSPlayStop)(void);
BOOL (STTSWINAPI *STTSQueryPlayStatus)(int *nStatus);
BOOL (STTSWINAPI *STTSAbout)(int nAboutType,TTSCHAR* pszAboutInfo, int* pnInfoSize);

//////////////////////////////////////////////////////////////////////////////////////////////
////Load TTS library,and get the procedure address of the dll

BOOL STTSLoadLibrary()
{
	DWORD dwErr;

	hSTTSApiMod=LoadLibrary(_T("sttsapi.dll"));
	if(hSTTSApiMod==NULL) {
		dwErr=GetLastError();
		SetLastError(dwErr);
		bSTTSApiLoaded = FALSE;
		return bSTTSApiLoaded;
	}

	STTSInit = (BOOL (STTSWINAPI *)())GetProcAddress(hSTTSApiMod,"STTSInit");
	STTSDeinit = (BOOL (STTSWINAPI *)())GetProcAddress(hSTTSApiMod,"STTSDeinit");
	STTSConnect = (HTTSINSTANCE (STTSWINAPI *)(TTSCHAR* ,TTSCHAR*))GetProcAddress(hSTTSApiMod,"STTSConnect");
	STTSDisconnect = (BOOL (STTSWINAPI *)(HTTSINSTANCE))GetProcAddress(hSTTSApiMod,"STTSDisconnect");
	STTSSetParam = (BOOL (STTSWINAPI *)(HTTSINSTANCE,int,int))GetProcAddress(hSTTSApiMod,"STTSSetParam");
	STTSGetParam = (BOOL (STTSWINAPI *)(HTTSINSTANCE,int,int*))GetProcAddress(hSTTSApiMod,"STTSGetParam");
	STTSString2AudioFile = (BOOL (STTSWINAPI *)(HTTSINSTANCE,TTSCHAR*,TTSCHAR*))GetProcAddress(hSTTSApiMod,"STTSString2AudioFile");
	STTSFile2AudioFile = (BOOL (STTSWINAPI *)(HTTSINSTANCE,TTSCHAR*,TTSCHAR*))GetProcAddress(hSTTSApiMod,"STTSFile2AudioFile");
	STTSPlayString = (BOOL (STTSWINAPI *)(HTTSINSTANCE,TTSCHAR*,BOOL))GetProcAddress(hSTTSApiMod,"STTSPlayString");
	STTSPlayTextFile = (BOOL (STTSWINAPI *)(HTTSINSTANCE,TTSCHAR*,BOOL))GetProcAddress(hSTTSApiMod,"STTSPlayTextFile");
	STTSPlayStop = (BOOL (STTSWINAPI *)(void))GetProcAddress(hSTTSApiMod,"STTSPlayStop");
	STTSQueryPlayStatus = (BOOL (STTSWINAPI *)(int*))GetProcAddress(hSTTSApiMod,"STTSQueryPlayStatus");
	STTSAbout = (BOOL (STTSWINAPI *)(int,TTSCHAR*,int* ))GetProcAddress(hSTTSApiMod,"STTSAbout");

	bSTTSApiLoaded=TRUE;

	return bSTTSApiLoaded;
}

///////////////////////////////////////////////////////////////////////////////////////////////
////Unload TTS library
BOOL STTSUnloadLibrary()
{

	DWORD dwErr;
	if(hSTTSApiMod!=NULL) {
		if(!FreeLibrary(hSTTSApiMod)) {
			dwErr=GetLastError();
			SetLastError(dwErr);
			return FALSE;
		}

		hSTTSApiMod=NULL;
		bSTTSApiLoaded = FALSE;
	}

	return TRUE;
}
