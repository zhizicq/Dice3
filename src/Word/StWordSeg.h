// StWordSeg.h: interface for the CStWordSeg class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_STWORDSEG_H__F0C5BE82_A7C2_4E88_A499_504AA2ACFF6E__INCLUDED_)
#define AFX_STWORDSEG_H__F0C5BE82_A7C2_4E88_A499_504AA2ACFF6E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
#include "afxtempl.h"
/*Word文档中的试题信息处理类*/
class CStWordSeg;
//typedef CMap<int,int, CStWordSeg*,CStWordSeg*>CStDocMap;
typedef CTypedPtrList<CPtrList,CStWordSeg*> CStDocMap;

class CStWordSeg 
{
public:
	CStWordSeg();
	virtual ~CStWordSeg();
public:
	static CStWordSeg* Get_StWordSeg_Obj(CString Keyword);
	void Serialize(CArchive &ar);
	static void FreeStWordSegObj(CStDocMap* pObjList);
	static void Serialize_Tab(CArchive& ar,CStDocMap* pObjList);
	void AddParamStartPos(int iStart, int Page, int Page_Line, CString &St_Param, CPtrList *pParamTab);
	void AddParamLine(int Line, int Page,int Page_Line,CString& St_Param,CPtrList* pParamTab=NULL);
	static CStWordSeg* FindNextObj(CStWordSeg* pCurObj);
	static CStWordSeg* FindKeyWordObj(int iKeyword,char* sKeyWord);
	void FreeMapTab();
	void AddParamLine(int Line,CString& St_Param);
	int m_iStart;//试题在文档的开始位置

	int m_Page;//页号,1-
	int m_Page_Line;//页内行号,1-
	int m_Para_Line;//参数行
	int m_keyword_Line;//知识点行
	int m_Tigan_Line;//题干行
	int m_Daan_Line;//答案行
	static int m_Doc_Max_Line;//Word文档的最大行号
	CString m_Tx;//题型代码
	CString m_Nd;//难度代码
	CString m_Zj;//章节代码
	double  m_Fz;//分值
	static CStDocMap m_DocSts;//映射表
};

#endif // !defined(AFX_STWORDSEG_H__F0C5BE82_A7C2_4E88_A499_504AA2ACFF6E__INCLUDED_)
