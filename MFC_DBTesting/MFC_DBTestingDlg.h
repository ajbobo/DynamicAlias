// MFC_DBTestingDlg.h : header file
//

#pragma once
#include "afxdb.h"
#include "afxwin.h"
#include "afxcmn.h"
#include "../DynamicAlias.h"

// CMFC_DBTestingDlg dialog
class CMFC_DBTestingDlg : public CDialog
{
// Construction
public:
	CMFC_DBTestingDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	enum { IDD = IDD_MFC_DBTESTING_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOpen();
	afx_msg void OnBnClickedClose();
	afx_msg void OnBnClickedExecuteSql();
private:
	CDatabase TheDB;
public:
	CString SQL_Query;
	CButton OpenButton;
	CButton CloseButton;
	CButton SQLButton;
	CEdit SQLEdit;
	CEdit DSNEdit;
	CListCtrl ResultList;
	afx_msg void OnLvnItemchangedResultlist(NMHDR *pNMHDR, LRESULT *pResult);
};
