//***************************************************************************************/
// File			:STTSApi.h
// Author		:zzh
// Aate			:2001/2/27
// Last update	:2001/3/26
// Company		:iFLYTEK
// Remarks		:This head file defines the interfaces of the simple TTS service.
// ***************************************************************************************/

#ifndef __STTSAPI_H__
#define __STTSAPI_H__

#include "iFly_TTS.h"

/////////////////////////////////////////////////////////////////////////////////////////////
#ifdef __cplusplus
extern "C" {
#endif

#ifdef  WIN32
#define STTSWINAPI WINAPI
#else
#define STTSWINAPI
#endif 

////////////////////////////////////////////////////////////////////////////////////////////
////Constans

#define STTS_TEXT_FILE			1
#define STTS_TEXT_STRING		2


////////////////////////////////////////////////////////////////////////////////////////////
////STTS SDK Functions

//------------------------------------------------------------------------------------------
// Function name: STTSInit
// Return value	: TRUE if the function is successful; otherwise FALSE.
// Parameters	:void
// Remarks		: Initialize the  STTS (simple TTS),it must be the first funtion to be called.
//------------------------------------------------------------------------------------------- 
BOOL STTSWINAPI STTSInit();

//-------------------------------------------------------------------------------------------
// Function name: STTSDeinit
// Return value	:TRUE if the function is successful; otherwise FALSE.
// Parameters	:void
// Remarks		:Unload the TTS from memory,it must be the last function to be called.
//-------------------------------------------------------------------------------------------
BOOL STTSWINAPI STTSDeinit();

//-------------------------------------------------------------------------------------------
// Function name:STTSConnect
// Return value	:The handle instance of the TTS service if the function is successfull,
//			NULL if connect to TTS server failed.
// Parameters	:
//		pszSerialNumber[in]:
//			The serial number of the TTS product which provided by USTC
//          IFLYTEK.CO. to initialize your TTS
//		pszTTSServerIP[in]:
//			IP address of the TTS Server to process net TTS. If the value is NULL,
//			It will process local TTS or automatically select a TTS server IP when
//			you have installed the dynamic load balance to your TTS server.
// Remarks		:It must connect to TTS server before use TTS service,and save the handle  
//			instance so that to use it next time.
//-------------------------------------------------------------------------------------------
HTTSINSTANCE STTSWINAPI STTSConnect(TTSCHAR* pszSerialNumber,TTSCHAR* pszTTSServerIP );
//void * (WINAPI *STTSConnect)(TTSCHAR* pszSerialNumber,TTSCHAR* pszTTSServerIP );
//-------------------------------------------------------------------------------------------
// Function name:STTSDisconnect
// Return value:TRUE if the function is successful; otherwise FALSE.
// Parameters:
//		hTTSInstance[in]:
//			The handle instance of the TTS service
// Remarks:Disconnect to the TTS system and destroy the hanle instance of TTS service .
//-------------------------------------------------------------------------------------------
BOOL STTSWINAPI STTSDisconnect(HTTSINSTANCE hTTSInstance);

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
BOOL STTSWINAPI STTSSetParam(HTTSINSTANCE hTTSInstance,int nType,int nParam);

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
BOOL STTSWINAPI STTSGetParam(HTTSINSTANCE hTTSInstance,int nType, int* nParam);

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
BOOL STTSWINAPI STTSFile2AudioFile(HTTSINSTANCE hTTSInstance,TTSCHAR* pszTextFile, TTSCHAR* pszAudioFile);

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
BOOL STTSWINAPI STTSString2AudioFile(HTTSINSTANCE hTTSInstance, TTSCHAR* pszString, TTSCHAR* pszAudioFile);

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
BOOL STTSWINAPI STTSPlayString(HTTSINSTANCE hTTSInstance , TTSCHAR* pszString,BOOL bASynch);

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
BOOL STTSWINAPI STTSPlayTextFile(HTTSINSTANCE hTTSInstance ,TTSCHAR* pszTextFile,BOOL bASynch);

//-------------------------------------------------------------------------------------------
// Function name:STTSPlayStop
// Return value:TRUE if the function is successful; otherwise FALSE.
// Parameters: void
// Remarks:Stop playing an audio.
//-------------------------------------------------------------------------------------------
BOOL STTSWINAPI STTSPlayStop(void );

//-------------------------------------------------------------------------------------------
// Function name:STTSQueryPlayStatus
// Return value:TRUE if the function is successful; otherwise FALSE.
// Parameters:
//		nStatus[out]: The status of playing,1 for playing,0 for else.
// Remarks:Query whether playing has finished on asynchronism playing mode
//-------------------------------------------------------------------------------------------
BOOL STTSWINAPI STTSQueryPlayStatus(int* nStatus );

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
BOOL STTSWINAPI STTSAbout(int nAboutType,TTSCHAR* pszAboutInfo, int* pnInfoSize);

#ifdef __cplusplus
} //extern "C" 
#endif

#endif

////////////////////////////////////////////////////////////////////////////////////////////
////End of STTSApi.h