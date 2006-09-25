// DynamicAlias.cpp : Defines the entry point for the DLL application.
//

#include "stdafx.h"
#include "stdio.h"
#include "odbcinst.h"
#include "DynamicAlias.h"

//#define DLL_EXPORT extern "C" __declspec(dllexport)
#define MAX_STRING_LEN 255

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
    return TRUE;
}

int CreateAlias(char *AliasName, char *filename, char *filetype)
{
	char dsndata[MAX_STRING_LEN];
	char *tempptr = dsndata;
	int val = 0;

	val = sprintf(tempptr,"DSN=%s",AliasName);
	tempptr = dsndata + val;
	dsndata[val] = 0;
	tempptr++;
	val = sprintf(tempptr,"DBQ=%s\0",filename);

	BOOL res = SQLConfigDataSource(NULL,ODBC_ADD_DSN,filetype,dsndata);

	return res;
}

int RemoveAlias(char *AliasName, char *filetype)
{
	char dsndata[MAX_STRING_LEN];

	sprintf(dsndata,"DSN=%s",AliasName);

	BOOL res = SQLConfigDataSource(NULL,ODBC_REMOVE_DSN,filetype,dsndata);

	return res;
}
