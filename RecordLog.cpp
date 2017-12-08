// RecordLog.cpp: implementation of the CRecordLog class.
//
//////////////////////////////////////////////////////////////////////
#include "stdafx.h"
#include "RecordLog.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////
CRecordLog *const CRecordLog::m_pInstance = NewRecordLog();
FILE *CRecordLog::pfLog = NULL;
CRITICAL_SECTION CRecordLog::m_csLog;

CRecordLog::CRecordLog()
{
	IntiInstance();
}

CRecordLog::~CRecordLog()
{
	if (pfLog != NULL)
	{
		fclose(pfLog);
		pfLog = NULL;
	}
	
	DeleteCriticalSection(&m_csLog);
}

BOOL CRecordLog::OpenLogFile()
{
	char strOldTime[20] = {0};
	char strBuffer[128] = {0};
	SYSTEMTIME SysTime;
	char strPath[256] = {0};
	char strCurTime[20];
	
	//取得当前目录
	char strFileName[256] = {0};

	GetSystemDirectory(strFileName, 128);
	strcat(strFileName, "\\CallInfo.ini");
	GetPrivateProfileString("LOG", "LOGPATH", "C:\\", strBuffer, 64, strFileName);
	strcpy(strPath, strBuffer);
	strcat(strPath, "\\ttsclient");
	mkdir(strPath);
	//取得系统时间
	GetLocalTime(&SysTime);
	sprintf(strCurTime, "%d-%2.2d-%2.2d", SysTime.wYear, SysTime.wMonth, SysTime.wDay);
	if(strcmp(strOldTime, strCurTime) != 0)
	{
		if(pfLog != NULL)
			fclose(pfLog);

		strcat(strPath, "\\");
		strcat(strPath, strCurTime);
		strcat(strPath, ".log");
		strcpy(strOldTime, strCurTime);

		if((pfLog = fopen(strPath, "a+")) == NULL)
			return FALSE;
		else
			return TRUE;
	}

	return TRUE;
}

//初始化成员变量
BOOL CRecordLog::IntiInstance()
{
	InitializeCriticalSection(&m_csLog);  //初化临界对象

	return OpenLogFile();
}

void CRecordLog::WriteLog(const char *pstrLog, UINT Len)
{
	SYSTEMTIME SysTime;
	char strTime[30] = {0};
	char strLog[LOG_BUF_SIZE + 20];

	if (pstrLog == NULL || Len <= 0)
		return;


	EnterCriticalSection(&m_csLog);
	
	//取得系统时间
	GetLocalTime(&SysTime);
	sprintf(strTime, "%d/%2.2d/%2.2d %d:%2.2d:%2.2d", SysTime.wYear, SysTime.wMonth, 
	SysTime.wDay, SysTime.wHour, SysTime.wMinute, SysTime.wSecond);
	
	//保存日志
	strcpy(strLog, strTime); 
	strcat(strLog, pstrLog);
	fwrite(strLog, 1, strlen(strLog), pfLog); 
	fflush(pfLog);

	OpenLogFile();

	LeaveCriticalSection(&m_csLog);
}