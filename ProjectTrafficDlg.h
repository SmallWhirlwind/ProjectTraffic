// ProjectTrafficDlg.h : header file
//

#if !defined(AFX_PROJECTTRAFFICDLG_H__7D12ADB9_23E7_49F5_80F1_659F0F65F418__INCLUDED_)
#define AFX_PROJECTTRAFFICDLG_H__7D12ADB9_23E7_49F5_80F1_659F0F65F418__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/////////////////////////////////////////////////////////////////////////////
// CProjectTrafficDlg dialog

class CProjectTrafficDlg : public CDialog
{
// Construction
public:
	CProjectTrafficDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	//{{AFX_DATA(CProjectTrafficDlg)
	enum { IDD = IDD_PROJECTTRAFFIC_DIALOG };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CProjectTrafficDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CProjectTrafficDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnButtonRead();
	afx_msg void OnButtonWrite();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PROJECTTRAFFICDLG_H__7D12ADB9_23E7_49F5_80F1_659F0F65F418__INCLUDED_)
