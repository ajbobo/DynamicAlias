// MFC_DBTesting.h : main header file for the PROJECT_NAME application
//

#pragma once

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols


// CMFC_DBTestingApp:
// See MFC_DBTesting.cpp for the implementation of this class
//

class CMFC_DBTestingApp : public CWinApp
{
public:
	CMFC_DBTestingApp();

// Overrides
	public:
	virtual BOOL InitInstance();

// Implementation

	DECLARE_MESSAGE_MAP()
};

extern CMFC_DBTestingApp theApp;