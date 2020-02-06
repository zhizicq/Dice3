// StWordSeg.cpp: implementation of the CStWordSeg class.
//
//////////////////////////////////////////////////////////////////////
#include "pch.h"
#include "stdafx.h"
#include "StWordSeg.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////
int CStWordSeg::m_Doc_Max_Line=0;//Word�ĵ�������к�
CStDocMap CStWordSeg::m_DocSts;
CStWordSeg::CStWordSeg()
{
	m_Para_Line=0;//������
	m_keyword_Line=0;//֪ʶ����
	m_Tigan_Line=0;//�����
	m_Daan_Line=0;//����
	m_Fz=0.0;
	m_Tx.Empty();
}

CStWordSeg::~CStWordSeg()
{
	FreeMapTab();
}
//��ӳ�������Ӳ�������Ϣ
void CStWordSeg::AddParamLine(int Line, CString& St_Param)
{
	if(Line<1)
		return;
	CStWordSeg* pLineObj=new CStWordSeg;
	pLineObj->m_Para_Line=Line;
	pLineObj->m_Tx=St_Param;
//	pLineObj->m_Nd=Nd;
//	pLineObj->m_Zj=Zj;
	//����ӳ���
	m_DocSts.AddTail(pLineObj);
}

void CStWordSeg::FreeMapTab()
{
	POSITION pos=m_DocSts.GetHeadPosition();
	while(pos)
	{
		CStWordSeg *pWordSeg;
		pWordSeg=m_DocSts.GetNext(pos);
		delete pWordSeg;
	}
	m_DocSts.RemoveAll();
}



CStWordSeg* CStWordSeg::FindKeyWordObj(int iKeyword, char* sKeyWord)
{
	POSITION pos = m_DocSts.GetHeadPosition();
	while (pos)
	{
		CStWordSeg* pObj = m_DocSts.GetNext(pos);
		if (pObj)
		{
			if (pObj->m_Para_Line == iKeyword)
			{
				sprintf(sKeyWord, "%s", "����");
				return pObj;//������
			}
			else if (pObj->m_keyword_Line == iKeyword)
			{
				sprintf(sKeyWord, "%s", "֪ʶ��");
				return pObj;//֪ʶ����
			}
			else if (pObj->m_Tigan_Line == iKeyword)
			{
				sprintf(sKeyWord, "%s", "�������");
				return pObj;//�����
			}
			else if (pObj->m_Daan_Line == iKeyword)
			{
				sprintf(sKeyWord, "%s", "��");
				return pObj;//����
			}
		}

	}
	return NULL;

}

CStWordSeg* CStWordSeg::FindNextObj(CStWordSeg *pCurObj)
{
	POSITION pos=m_DocSts.Find(pCurObj);
	if(pos)
	{
		m_DocSts.GetNext(pos);
		if(pos)
			return m_DocSts.GetNext(pos);
	}
	return NULL;
}






void CStWordSeg::AddParamLine(int Line, int Page, int Page_Line, CString &St_Param,CPtrList* pParamTab)
{
	if(Line<1||pParamTab==NULL)
		return;
	CStWordSeg* pLineObj=new CStWordSeg;
	pLineObj->m_Page=Page;
	pLineObj->m_Page_Line=Page_Line;
	pLineObj->m_Para_Line=Line;
	int idx=St_Param.Find(':');
	pLineObj->m_Tx=St_Param.Left(idx);
	CString str=St_Param.Mid(idx+1);
	double fz=0.0;
	sscanf(str,"%lf",&fz);
	if(fz>0)
		pLineObj->m_Fz=fz;
	//����ӳ���
	pParamTab->AddTail(pLineObj);
}

void CStWordSeg::AddParamStartPos(int iStart, int Page, int Page_Line, CString &St_Param, CPtrList *pParamTab)
{
	if(iStart<1||pParamTab==NULL)
		return;
	CStWordSeg* pLineObj=new CStWordSeg;
	pLineObj->m_Page=Page;//ҳ��
	pLineObj->m_Page_Line=Page_Line;//ҳ���к�
	pLineObj->m_iStart=iStart;
	pLineObj->m_Para_Line=Page_Line;//�к�
	int idx=St_Param.Find(':');
	//ȡ���
	pLineObj->m_Tx=St_Param.Left(idx);
	//ȡ��ֵ
	CString str=St_Param.Mid(idx+1);
	double fz=0.0;
	sscanf(str,"%lf",&fz);
	if(fz>0)
		pLineObj->m_Fz=fz;
	//����ӳ���
	pParamTab->AddTail(pLineObj);
}

void CStWordSeg::Serialize_Tab(CArchive &ar, CStDocMap *pObjList)
{
	if(pObjList==NULL)
		return;
	if(ar.IsStoring())
	{
		int ncs=pObjList->GetCount();
		ar<<ncs;
		POSITION pos=pObjList->GetHeadPosition();
		while(pos)
		{
			CStWordSeg* pObj=pObjList->GetNext(pos);
			pObj->Serialize(ar);
		}
	}
	else
	{
		FreeStWordSegObj(pObjList);
		int ncs=0;
		ar>>ncs;
		for(;ncs>0;ncs--)
		{
			CStWordSeg* pObj=new CStWordSeg;
			pObj->Serialize(ar);
			pObjList->AddTail(pObj);
		}
	}
}

void CStWordSeg::Serialize(CArchive &ar)
{
	if(ar.IsStoring())
	{
		ar<<m_Page;
		ar<<m_Page_Line;//ҳ���к�,1-
		ar<<m_Para_Line;//������
		ar<<m_Tx;//���ʹ���
		ar<<m_Fz;//��ֵ
	}
	else
	{
		ar>>m_Page;
		ar>>m_Page_Line;//ҳ���к�,1-
		ar>>m_Para_Line;//������
		ar>>m_Tx;//���ʹ���
		ar>>m_Fz;//��ֵ
	}
}

void CStWordSeg::FreeStWordSegObj(CStDocMap *pObjList)
{
	if(pObjList)
	{
		POSITION pos=pObjList->GetHeadPosition();
		while(pos)
		{
			CStWordSeg* pObj=pObjList->GetNext(pos);
			if(pObj)
				delete pObj;
		}
		pObjList->RemoveAll();
	}
}

CStWordSeg* CStWordSeg::Get_StWordSeg_Obj(CString Keyword)
{
	
	POSITION pos=m_DocSts.GetHeadPosition();
	while(pos)
	{
		CStWordSeg *pObj=m_DocSts.GetNext(pos);
		if(pObj && pObj->m_Tx.CompareNoCase(Keyword)==0)
			return pObj;
	}
	return NULL;
}
