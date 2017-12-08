// TTSClient.cpp : Defines the entry point for the DLL application.
//

#include "stdafx.h"
#include "TTSClient.h"
#include "iflysttsapi.h"

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
	BOOL    bSTTSInit, bSTTSApiLoaded;
	// Perform actions based on the reason for calling.
    switch(ul_reason_for_call) 
    { 
    case DLL_PROCESS_ATTACH:
        // Initialize once for each new process.
        // Return FALSE to fail DLL load.
		//装载STTSApi.dll和iFlyTTS.dll
		bSTTSApiLoaded = STTSLoadLibrary();
		if(!bSTTSApiLoaded)
			return FALSE;
		bSTTSInit = STTSInit();
		if(!bSTTSInit)
			return FALSE;
        break;
        
    case DLL_THREAD_ATTACH:
        // Do thread-specific initialization.
		//装载STTSApi.dll和iFlyTTS.dll
		bSTTSApiLoaded = STTSLoadLibrary();
		if(!bSTTSApiLoaded)
			return FALSE;
		bSTTSInit = STTSInit();
		if(!bSTTSInit)
			return FALSE;
        break;
        
    case DLL_THREAD_DETACH:
        // Do thread-specific cleanup.
		//合成完毕,卸载STTSApi.dll和iFlyTTS.dll
		bSTTSApiLoaded = STTSUnloadLibrary();
		if(!bSTTSApiLoaded)
		{
			WRITELOG("unload FlyTTS library  failure !");
			return 0;
		}
        break;
        
    case DLL_PROCESS_DETACH:
        // Perform any necessary cleanup.
		//合成完毕,卸载STTSApi.dll和iFlyTTS.dll
		bSTTSApiLoaded = STTSUnloadLibrary();
		if(!bSTTSApiLoaded)
		{
			WRITELOG("unload FlyTTS library  failure !");
			return 0;
		}
        break;
    }

    return TRUE;
}

int String2AudioFile(char* psString, char* psAudioFile)
{
	int		nLen;
	int     nSpeed;
	int     nAudioFmt = TTS_ADF_ALAW8K1C;
	int     nCodePage = 1;
	char    sBuf[30] = {0};
	char    sServer[16] = {0};
	HTTSINSTANCE hTTSInstance;
	//BOOL    bSTTSInit, bSTTSApiLoaded;
	char	sString[STRING_SIZE];
	char	sAudioFile[FILENAME_SIZE];
	char    sSerialNo[]="*****-*****-*****";

	nLen = strlen(psString);
	if(nLen > STRING_SIZE)
	{
		memcpy(sString, psString, STRING_SIZE - 2);
		sString[STRING_SIZE - 1] = '\0';
	}
	else
	{
		memcpy(sString, psString, nLen);
		sString[nLen] = '\0';
	}
    
	nLen = strlen(psAudioFile);
	if(nLen > STRING_SIZE)
	{
		memcpy(sAudioFile, psAudioFile, FILENAME_SIZE - 2);
		sAudioFile[FILENAME_SIZE - 1] = '\0';
	}
	else
	{
		memcpy(sAudioFile, psAudioFile, nLen);
		sAudioFile[nLen] = '\0';
	}	

	GetPrivateProfileString("SYSTEM", "TTSSERVER", NULL, sBuf, 15, ".\\para.ini");
	strcpy(sServer, sBuf);
	nSpeed = GetPrivateProfileInt("SYSTEM", "SYNTHSPEED", 0, ".\\para.ini");
	
	WRITELOG(sSerialNo);
	WRITELOG(sServer);
	
	if(!(hTTSInstance = STTSConnect(sSerialNo, sServer)))
	{
		WRITELOG("connected tts server failur !");
        return 0;
	}

	//设置合成参数，包括内码，音频格式和合成语速
	if(!STTSSetParam(hTTSInstance, TTS_PARAM_CODEPAGE, nCodePage)) 
	{
		WRITELOG("set CodePage error then use default internal CodePage !");
		STTSDisconnect(hTTSInstance);
		return 0;
	}	

	if(!STTSSetParam(hTTSInstance, TTS_PARAM_AUDIODATAFMT, nAudioFmt))
	{
		DWORD dwErr = GetLastError();
		if(TTSGETERRCODE(dwErr) == TTSERR_OK )
		{
			WRITELOG("executed successe !");
		}
		else if(TTSGETERRCODE(dwErr) == TTSERR_INVALIDHANDLE)
		{
			WRITELOG("invalid TTS instance !");
		}
		else if(TTSGETERRCODE(dwErr) == TTSERR_NOTSUPP)
		{
			WRITELOG("system nonsupport this operated !");
		}
		else if(TTSGETERRCODE(dwErr) == TTSERR_INVALIDPARA)
		{
			WRITELOG("parameter error !");
		}
		else if(TTSGETERRCODE(dwErr) == TTSERR_LOADDLL)
		{
			WRITELOG("load FlyTTS library  failure !");
		}
		STTSDisconnect(hTTSInstance);
		return 0;
	}

	if(!STTSSetParam(hTTSInstance, TTS_PARAM_SPEED, nSpeed))
	{
		WRITELOG("set speed error then use default value !");
		STTSDisconnect(hTTSInstance);
        return 0;
	}	
	
	//合成字符串，调用合成字符串的接口
	if(!STTSString2AudioFile(hTTSInstance,psString,psAudioFile))
	{
		DWORD dwErr = GetLastError();
		WRITELOG("string Transfer to audioFile failure!");
		STTSDisconnect(hTTSInstance);
		return 0;
	}	

	WRITELOG(sAudioFile);
	WRITELOG("string Transfer to audioFile success!");		
	STTSDisconnect(hTTSInstance);

	return 1;
}

int TextFile2AudioFile(char* psTextFile, char* psAudioFile)
{
	int		nLen;
	int     nSpeed;
	int     nAudioFmt = TTS_ADF_ALAW8K1C;
	int     nCodePage = 1;
	char    sBuf[30] = {0};
	char    sServer[16] = {0};
	HTTSINSTANCE hTTSInstance;
	//BOOL    bSTTSInit, bSTTSApiLoaded;
	char	sTextFile[STRING_SIZE];
	char	sAudioFile[FILENAME_SIZE];
	char    sSerialNo[]="*****-*****-*****";

	nLen = strlen(psTextFile);
	if(nLen > STRING_SIZE)
	{
		memcpy(sTextFile, psTextFile, STRING_SIZE - 2);
		sTextFile[STRING_SIZE - 1] = '\0';
	}
	else
	{
		memcpy(sTextFile, psTextFile, nLen);
		sTextFile[nLen] = '\0';
	}
    
	nLen = strlen(psAudioFile);
	if(nLen > STRING_SIZE)
	{
		memcpy(sAudioFile, psAudioFile, FILENAME_SIZE - 2);
		sAudioFile[FILENAME_SIZE - 1] = '\0';
	}
	else
	{
		memcpy(sAudioFile, psAudioFile, nLen);
		sAudioFile[nLen] = '\0';
	}	

	GetPrivateProfileString("SYSTEM", "TTSSERVER", NULL, sBuf, 15, ".\\para.ini");
	strcpy(sServer, sBuf);
	nSpeed = GetPrivateProfileInt("SYSTEM", "SYNTHSPEED", 0, ".\\para.ini");
	
	WRITELOG(sSerialNo);
	WRITELOG(sServer);
	
	if(!(hTTSInstance = STTSConnect(sSerialNo, sServer)))
	{
		WRITELOG("connected tts server failur !");
        return 0;
	}

	//设置合成参数，包括内码，音频格式和合成语速
	if(!STTSSetParam(hTTSInstance, TTS_PARAM_CODEPAGE, nCodePage)) 
	{
		WRITELOG("set CodePage error then use default internal CodePage !");
		STTSDisconnect(hTTSInstance);
		return 0;
	}	

	if(!STTSSetParam(hTTSInstance, TTS_PARAM_AUDIODATAFMT, nAudioFmt))
	{
		DWORD dwErr = GetLastError();
		if(TTSGETERRCODE(dwErr) == TTSERR_OK )
		{
			WRITELOG("executed successe !");
		}
		else if(TTSGETERRCODE(dwErr) == TTSERR_INVALIDHANDLE)
		{
			WRITELOG("invalid TTS instance !");
		}
		else if(TTSGETERRCODE(dwErr) == TTSERR_NOTSUPP)
		{
			WRITELOG("system nonsupport this operated !");
		}
		else if(TTSGETERRCODE(dwErr) == TTSERR_INVALIDPARA)
		{
			WRITELOG("parameter error !");
		}
		else if(TTSGETERRCODE(dwErr) == TTSERR_LOADDLL)
		{
			WRITELOG("load FlyTTS library  failure !");
		}
		STTSDisconnect(hTTSInstance);
		return 0;
	}

	if(!STTSSetParam(hTTSInstance, TTS_PARAM_SPEED, nSpeed))
	{
		WRITELOG("set speed error then use default value !");
		STTSDisconnect(hTTSInstance);
        return 0;
	}	
	
	//合成字符串，调用合成字符串的接口
	if(!STTSFile2AudioFile(hTTSInstance,psTextFile,psAudioFile))
	{
		DWORD dwErr = GetLastError();
		WRITELOG("Text File Transfer to audioFile failure!");
		STTSDisconnect(hTTSInstance);
		return 0;
	}	

	WRITELOG(sAudioFile);
	WRITELOG("Text File Transfer to audioFile success!");		
	STTSDisconnect(hTTSInstance);

	return 1;
}


