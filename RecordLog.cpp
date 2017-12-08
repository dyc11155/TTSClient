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
	
	//ȡ�õ�ǰĿ¼
	char strFileName[256] = {0};

	GetSystemDirectory(strFileName, 128);
	strcat(strFileName, "\\CallInfo.ini");
	GetPrivateProfileString("LOG", "LOGPATH", "C:\\", strBuffer, 64, strFileName);
	strcpy(strPath, strBuffer);
	strcat(strPath, "\\ttsclient");
	mkdir(strPath);
	//ȡ��ϵͳʱ��
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

//��ʼ����Ա����
BOOL CRecordLog::IntiInstance()
{
	InitializeCriticalSection(&m_csLog);  //�����ٽ����

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
	
	//ȡ��ϵͳʱ��
	GetLocalTime(&SysTime);
	sprintf(strTime, "%d/%2.2d/%2.2d %d:%2.2d:%2.2d", SysTime.wYear, SysTime.wMonth, 
	SysTime.wDay, SysTime.wHour, SysTime.wMinute, SysTime.wSecond);
	
	//������־
	strcpy(strLog, strTime); 
	strcat(strLog, pstrLog);
	fwrite(strLog, 1, strlen(strLog), pfLog); 
	fflush(pfLog);

	OpenLogFile();

	LeaveCriticalSection(&m_csLog);
}