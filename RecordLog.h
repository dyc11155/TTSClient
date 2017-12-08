// RecordLog.h: interface for the CRecordLog class.
//
//////////////////////////////////////////////////////////////////////

#ifndef RECORD_LOG_H
#define RECORD_LOG_H

#include <windows.h>
#include <stdio.h>
#include <assert.h>
#include <string.h>
#include <winbase.h>
#include <direct.h>

#define LOG_BUF_SIZE  2000 

class CRecordLog
{
private:
	CRecordLog();

public:
	virtual ~CRecordLog();
				
public:
	static CRecordLog &GetInstance()
	{
		return *m_pInstance;
	}
	static void WriteLog(const char *pstrLog, UINT Len);

protected:
	static CRecordLog *NewRecordLog()
	{
		static CRecordLog RecordLog;
		return &RecordLog;
	}
	BOOL IntiInstance();
	static BOOL OpenLogFile();

protected:
	static CRecordLog *const	m_pInstance;
	static CRITICAL_SECTION		m_csLog;			//临界对象
	static FILE					*pfLog;				//日志文件指针 	
};

//#if _DEBUG 
	#define WRITELOG(exp) \
	{ \
		char Buffer[LOG_BUF_SIZE] = {0}; \
		_snprintf(Buffer, LOG_BUF_SIZE, " File %s Line %d %s\n", __FILE__, __LINE__, exp); \
		CRecordLog::WriteLog(Buffer, strlen(Buffer)); \
	}
//#else
	//#define WRITELOG(exp) {}
//#endif
/*(void)((exp) || (CRecordLog::WriteLog(pBuf, strlen(pBuf)), 0));*/
#endif
