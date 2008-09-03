#ifndef _DYNAMIC_ALIAS_H
#define _DYNAMIC_ALIAS_H

// Constants used all the time
#define EXCEL_ENGLISH "Microsoft Excel Driver (*.xls)"
#define ACCESS_ENGLISH "Microsoft Access Driver (*.mdb)"

// If DYNAMICALIAS_EXPORTS is defined (only in the DLL's original project) export data. If not, set up import features
#ifdef DYNAMICALIAS_EXPORTS

	#define DLL_EXPORT extern "C" __declspec(dllexport)

	DLL_EXPORT
	int CreateAlias(char *AliasName, char *filename, char *filetype);

	DLL_EXPORT
	int RemoveAlias(char *AliasName, char *filetype);

#else

	static int (*CreateAlias)(char *AliasName, char *filename, char *filetype);
	static int (*RemoveAlias)(char *AliasName, char *filetype);

	static HMODULE AliasLib;

	static int DynamicAlias_LoadFunctions()
	{
		HMODULE lib = LoadLibrary("DynamicAlias.DLL");
		AliasLib = lib;
		if (lib == NULL)
		{
			MessageBox(NULL,"Unable to load DLL","DLL Load Error",MB_OK);
			return 0;
		}
		if ((CreateAlias = (int(*)(char*,char*,char*))GetProcAddress(lib,"CreateAlias")) == NULL)
		{
			MessageBox(NULL,"Unable to load CreateAlias()","DLL Load Error",MB_OK);
			return 0;
		}
		if ((RemoveAlias = (int(*)(char*,char*))GetProcAddress(lib,"RemoveAlias")) == NULL)
		{
			MessageBox(NULL,"Unable to load RemoveAlias()","DLL Load Error",MB_OK);
			return 0;
		}
		return 1;
	}

	static int DynamicAlias_UnloadFunctions()
	{
		CreateAlias = NULL;
		RemoveAlias = NULL;
		return FreeLibrary(AliasLib);
	}

#endif

#endif