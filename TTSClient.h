#ifndef TTS_CLIENT_H 
#define TTS_CLIENT_H

#include <stdio.h>
#include "RecordLog.h"

#define STRING_SIZE 1000
#define FILENAME_SIZE 50

//以下等价于 extern "C" __declspec(dllexport) int String2AudioFile(char* psString, char* psAudioFile);
//用于开放一个名叫String2AudioFile的dll接口
#ifdef TTSCLIENT_EXPORTS
#define TTSCLIENT_API __declspec(dllexport)   
#else
#define TTSCLIENT_API __declspec(dllimport)  
#endif
	
#ifdef __cplusplus
	extern "C" 
	{
#endif

	TTSCLIENT_API  int String2AudioFile(char* psString, char* psAudioFile);
	TTSCLIENT_API  int TextFile2AudioFile(char* psTextFile, char* psAudioFile);

#ifdef __cplusplus
	}
#endif

#endif