// StWordSeg.h: interface for the CStWordSeg class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_STWORDSEG_H__F0C5BE82_A7C2_4E88_A499_504AA2ACFF6E__INCLUDED_)
#define AFX_STWORDSEG_H__F0C5BE82_A7C2_4E88_A499_504AA2ACFF6E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
#include "afxtempl.h"
/*Word�ĵ��е�������Ϣ������*/
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
	int m_iStart;//�������ĵ��Ŀ�ʼλ��

	int m_Page;//ҳ��,1-
	int m_Page_Line;//ҳ���к�,1-
	int m_Para_Line;//������
	int m_keyword_Line;//֪ʶ����
	int m_Tigan_Line;//�����
	int m_Daan_Line;//����
	static int m_Doc_Max_Line;//Word�ĵ�������к�
	CString m_Tx;//���ʹ���
	CString m_Nd;//�Ѷȴ���
	CString m_Zj;//�½ڴ���
	double  m_Fz;//��ֵ
	static CStDocMap m_DocSts;//ӳ���
};

#endif // !defined(AFX_STWORDSEG_H__F0C5BE82_A7C2_4E88_A499_504AA2ACFF6E__INCLUDED_)
