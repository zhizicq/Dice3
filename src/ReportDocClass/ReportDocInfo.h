// ReportDocInfo.h: interface for the CReportDocInfo class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_REPORTDOCINFO_H__D9BDD4D8_1055_4584_9CBC_6C7E666AA808__INCLUDED_)
#define AFX_REPORTDOCINFO_H__D9BDD4D8_1055_4584_9CBC_6C7E666AA808__INCLUDED_

#include "..\WORD\yzWordOperator.h"	// Added by ClassView
#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CReportDocInfo : public CObject  
{
public:
	void Check_TabCharPos();
	void Check_Bianhao();
	void Check_Tables();
	void Check_ShapePicture();
	void Check_WordShapes();
	void Check_Biaoti();
	void FindReplace(CString Keyword, CString Space,BOOL bWhile=TRUE);
	void Check_Paragraphs();
	void Check_PageSet();
	void Check_InlineShapes();
	CString GetKeywordInfo();
	void CheckKeywordItem();
	void GetInlineShapesInfo();
	void GetShapesInfo();
	void GetTableInfo();
	void GetParagraphs_Info();
	void GetPageSet();
	void SetCheck();

	void GetMuluInfo();
	CyzWordOperator m_WordApp;
	CReportDocInfo();
	virtual ~CReportDocInfo();
	double m_PageWide;//ҳ����cm
	double m_PageHigh;//ҳ��߶�cm
	double m_PageLeft;//ҳ����߾�
	double m_PageRight;//ҳ���ұ߾�
	double m_PageTop;//ҳ�涥�߾�
	double m_PageBottom;//ҳ��ױ߾�
	CString m_szError;
};

#endif // !defined(AFX_REPORTDOCINFO_H__D9BDD4D8_1055_4584_9CBC_6C7E666AA808__INCLUDED_)
