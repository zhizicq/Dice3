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
	double m_PageWide;//Ò³Ãæ¿í¶Ècm
	double m_PageHigh;//Ò³Ãæ¸ß¶Ècm
	double m_PageLeft;//Ò³Ãæ×ó±ß¾à
	double m_PageRight;//Ò³ÃæÓÒ±ß¾à
	double m_PageTop;//Ò³Ãæ¶¥±ß¾à
	double m_PageBottom;//Ò³Ãæµ×±ß¾à
	CString m_szError;
};

#endif // !defined(AFX_REPORTDOCINFO_H__D9BDD4D8_1055_4584_9CBC_6C7E666AA808__INCLUDED_)
