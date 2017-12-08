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
extern BOOL STTSLoadLibrary();

//------------------------------------------------------------------------------------------
// Function name: STTSunloadLibrary
// Return value : TRUE if the function is successful; otherwise FALSE.
// Remarks		: Unload STTSapi.dll
//------------------------------------------------------------------------------------------
extern BOOL  STTSUnloadLibrary();


////////////////////////////////////////////////////////////////////////////////////////////
////iFly STTS SDK functions

//-------------------------------------------------------------------------------------------
// Function name: STTSInit
// Return value: TRUE if the function is successful; otherwise FALSE.
// Parameters:void
// Remarks: Initialize the  STTS (simple TTS),it must be the first funtion to be called.
//------------------------------------------------------------------------------------------- 
extern BOOL (STTSWINAPI *STTSInit)();

//-------------------------------------------------------------------------------------------
// Function name: STTSDeinit
// Return value:TRUE if the function is successful; otherwise FALSE.
// Parameters:void
// Remarks:Unload the TTS from memory,it must be the last function to be called.
//-------------------------------------------------------------------------------------------

extern BOOL (STTSWINAPI *STTSDeinit)();
//-------------------------------------------------------------------------------------------
// Function name:STTSConnect
// Return value:The handle instance of the TTS service if the function is successfull,
//			NULL if connect to TTS server failed.
// Parameters:
//		pszSerialNumber[in]:
//			The serial number of the TTS product which provided by USTC
//          IFLYTEK.CO. to initialize your TTS
//		pszTTSServerIP[in]:
//			IP address of the TTS Server to process net TTS. If the value is NULL,
//			It will process local TTS or automatically select a TTS server IP when
//			you have installed the dynamic load balance to your TTS server.
// Remarks:It must connect to TTS server before use TTS service,and save the handle instance 
//			so that to use it next time.
//-------------------------------------------------------------------------------------------
extern HTTSINSTANCE (STTSWINAPI *STTSConnect)(TTSCHAR* pszSerialNumber,TTSCHAR* pszTTSServerIP);

//-------------------------------------------------------------------------------------------
// Function name:STTSDisconnect
// Return value:TRUE if the function is successful; otherwise FALSE.
// Parameters:
//		hTTSInstance[in]:
//			The handle instance of the TTS service
// Remarks:Disconnect to the TTS system and destroy the hanle instance of TTS service .
//-------------------------------------------------------------------------------------------
extern BOOL (STTSWINAPI *STTSDisconnect)(HTTSINSTANCE hTTSInstance);

//-------------------------------------------------------------------------------------------
// Function name:STTSSetParam
// Return value:TRUE if the function is successful; otherwise FALSE.
// Parameters:
//		hTTSInstance[in]:
//			The handle instance of the TTS service
//		nType[in]:
//			Parameter type.
//		nParam[in]:
//			parameter value.
// Remarks:Use this function to set TTS parameter value. 
//-------------------------------------------------------------------------------------------
extern BOOL (STTSWINAPI *STTSSetParam)(HTTSINSTANCE hTTSInstance,int nType,int nParam);

//-------------------------------------------------------------------------------------------
// Function name:STTSGetParam
// Return value:TRUE if the function is successful; otherwise FALSE.
// Parameters:
//		hTTSInstance[in]:
//			The handle instance of the TTS service
//		nType[in]:
//			Param type.
//		nParam[out]:
//			The address of the parameter's value	
// Remarks:Get the parameter's value.
//-------------------------------------------------------------------------------------------
extern BOOL (STTSWINAPI *STTSGetParam)(HTTSINSTANCE hTTSInstance,int nType, int* nParam);

//-------------------------------------------------------------------------------------------
// Function name:SString2AudioFile
// Return value:TRUE if the function is successful; otherwise FALSE.
// Parameters:
//		hTTSInstance[in]:
//			The handle instance of the TTS service
//		pszString[in]:
//			The string what to synthesize.
//		pszAudioFile[in]:
//			Audio file that hold the content after synthesis.
// Remarks:Make a string to an autio file using TTS.
//-------------------------------------------------------------------------------------------
extern BOOL (STTSWINAPI *STTSString2AudioFile)(HTTSINSTANCE hTTSInstance, TTSCHAR* pszString, TTSCHAR* pszAudioFile);

//-------------------------------------------------------------------------------------------
// Function name:SFile2AudioFile
// Return value:TRUE if the function is successful; otherwise FALSE.
// Parameters:
//		hTTSInstance[in]:
//			The handle instance of the TTS service
//		pszTextFile[in]:
//			Text file name which to be synthesized to audio file.
//		pszAudioFile[in]:
//			Audio file that hold the content after synthesis.
// Remarks:Convert a text file to an audio file take advantage of TTS.
//-------------------------------------------------------------------------------------------
extern BOOL (STTSWINAPI *STTSFile2AudioFile)(HTTSINSTANCE hTTSInstance,TTSCHAR* pszTextFile, TTSCHAR* pszAudioFile);

//-------------------------------------------------------------------------------------------
// Function name:STTSPlayString
// Return value:TRUE if the function is successful; otherwise FALSE.
// Parameters:
//		hTTSInstance[in]:
//			The handle instance of the TTS service
//		pszString[in]:
//			The string which to be synthesized.
//		bASynch[in]:
//			TRUE if playing on asynchronism mode,or FALSE on synchronism mode 
// Remarks:Synthesize a string to audio and play it.
//-------------------------------------------------------------------------------------------
extern BOOL (STTSWINAPI *STTSPlayString)(HTTSINSTANCE hTTSInstance , TTSCHAR* pszString,BOOL bASynch);

//-------------------------------------------------------------------------------------------
// Function name:STTSPlayTextFile
// Return value:TRUE if the function is successful; otherwise FALSE.
// Parameters:
//		hTTSInstance[in]:
//			The handle instance of the TTS service.
//		pszTextFile[in]:
//			Text file name to synthesize.
//		bASynch[in]:
//			TRUE if playing on asynchronism mode,or FALSE on synchronism mode 
// Remarks:Synthesize a text file to audio and play it.
//-------------------------------------------------------------------------------------------
extern BOOL (STTSWINAPI *STTSPlayTextFile)(HTTSINSTANCE hTTSInstance ,TTSCHAR* pszTextFile,BOOL bASynch);

//-------------------------------------------------------------------------------------------
// Function name:STTSPlayStop
// Return value:TRUE if the function is successful; otherwise FALSE.
// Parameters: void
// Remarks:Stop playing an audio.
//-------------------------------------------------------------------------------------------
extern BOOL (STTSWINAPI *STTSPlayStop)(void );

//-------------------------------------------------------------------------------------------
// Function name:STTSQueryPlayStatus
// Return value:TRUE if the function is successful; otherwise FALSE.
// Parameters:
//		nStatus[out]:1 for playing,otherwise 0.
// Remarks:Query whether playing has finished on asynchronism playing mode
//-------------------------------------------------------------------------------------------
extern BOOL (STTSWINAPI *STTSQueryPlayStatus)(int* nStatus);

//-------------------------------------------------------------------------------------------
// Function name:STTSAbout
// Return value:TRUE if the function is successful; otherwise FALSE.
// Parameters:
//		pszAboutInfo[in]:
//			The about information.
//		nInfoSize[in/out]:
//			Input the initial size of about information,and if the function successfull,
//			nInfoSize will output the really size of the information,otherwize the size  
//			that needed to hold the about information.  
// Remarks:The about information of the  TTS version.
//-------------------------------------------------------------------------------------------
extern BOOL (STTSWINAPI *STTSAbout)(int nAboutType,TTSCHAR* pszAboutInfo, int* pnInfoSize);

#endif