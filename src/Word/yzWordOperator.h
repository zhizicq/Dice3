// yzWordOperator.h: interface for the CyzWordOperator class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_YZWORDOPERATOR_H__C4696929_36D6_4DCC_9A81_91DBCF67A837__INCLUDED_)
#define AFX_YZWORDOPERATOR_H__C4696929_36D6_4DCC_9A81_91DBCF67A837__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
#include "afxtempl.h"
#include "StWordSeg.h"	// Added by ClassView
#include "msword.h"
#include <fstream>
using namespace WORD_NAMESPACE;
/*提取行数
a = ActiveDocument.BuiltInDocumentProperties(wdPropertyLines).Value
MsgBox Str(a)
*/
class CyzWordOperator
{
	#define MAX_DISP_ARGS 10
	#define DISPARG_NOFREEVARIANT 0x01
	#define DISP_FREEARGS 0x02
	#define DISP_NOSHOWEXCEPTIONS 0x03
public:
	BOOL GetCursorTable_Row_Col(long &TableNo, long &row, long &col);
	BOOL StartWord();
	void SetVisible(BOOL bVisible=FALSE);
	BOOL AddArgumentInt2(LPOLESTR lpszArgName, WORD wFlags, int i);
	BOOL CloseActiveDocument(BOOL bSaveIt);
	void ReleaseVariant(VARIANTARG *pvarg);
	void AddArgumentCommon(LPOLESTR lpszArgName, WORD wFlags, VARTYPE vt);
	BOOL AddArgumentCString(LPOLESTR lpszArgName, WORD wFlags, CString szStr);
	void ClearAllArgs();
	void ClearVariant(VARIANTARG *pvarg);
	BOOL OpenWordFile(CString szFileName);
	BOOL InsertBreakPage();
	void SetPageView(int model=3);
	CString Get_All_Text();
	BOOL StartWordApp();
	void All_Paste();
	void Paste();
	void WriteText(CString txt,WORD_NAMESPACE::_Application* pApp);
	void TypeTextCenter(CString txt,float size=0.0);
	void Insert_DocFile(CString szDocName);
	void Insert_TypeText(CString txt,CString szFontName,int Alignment=3,float size=0.0);
	void Insert_TypeText(CString txt,CString szFontName,int Alignment,float size,DWORD Color=0,BOOL bBold=FALSE);
	void SelectedCopy();
	void SetDisplayPageBoundaries(BOOL bVisable=FALSE);
	BOOL ProtectedDoc();
	void DispStParam(CListCtrl* pListCtrl);
	int GetPageLine(int& nLine);
	CString GetText();
	CString GetWords(int &chars);
	void GotoLine(int keyline);
	void DispResult(CListCtrl* pListCtrl);
	BOOL Find_KeyWord_Et(int StartLine,int EndLine,CStWordSeg* pObj=NULL);
	void Search_Keyword_et();
	CStWordSeg m_StInfoInDoc;
	BOOL FindKeyWord(CString& keyword);
	void SetSelectRangFont();
	void SetZoom(long bl=100);
	void SelectRange(int nStartLine,int nEndLine);
	int GetWordDocLines(BOOL bLine=TRUE);
	void ReportInfo();
	WORD_NAMESPACE::_Application* m_pWordApp;
	_Document m_WordDoc;
	_CommandBars m_WordCom;
	
	BOOL GetMaths(HWND hWnd, OMaths maths, InlineShapes isps);
	BOOL GetWordAppObj();
	Shape GetWordShapes();
	CyzWordOperator();
	CStDocMap* m_pStParamTab;//试题参数表
	int m_MaxPos;//文档的最后位置
	virtual ~CyzWordOperator();

	BOOL savePicture(HWND hWnd, CString FileName);
	
	WORD m_awFlags[MAX_DISP_ARGS];
public:
	BOOL GotoPageNo(int PageNo);
	void SetFontInfo(CString ZwName,CString XwName,float size,COLORREF color);
	long GetShapesGroupCount();
	//指定段首文字和结束文字，读取文本
	CString GetCurParagraphInfo(CString StartTxt,CString EndTxt);
	CString GetTableCellText(int iTabNo,int row,int col,long& endpos);
	int GetTableInfo();
	void ClearAll_HyperLink();
	BOOL ReplaceEquation(CString Keyword);
	void ShowEditFlag(BOOL bShow);
	void CopyItem(CWnd* pWnd,UINT uPaseCmd,BOOL& bEquation);
	void SetDisplayScreenTips(BOOL bShow=TRUE);
	CString GetHyperLinksText();
	void SetHyperLink(CString strLink);
	int GetSelectRangeKeyPos(CString keyword, int iStart, int iEnd);
	void SetDisplayRulers(BOOL bDisplay=TRUE);
	void SetMargin();
	BOOL Search_Char_Pos(CString find_Txt,int& Pos);
	CString GetKeywordLineInfo(CString Keyword);
	void Search_StInfo_Shjuan(CString Keyword, CStringArray &InfoAry);
	CString GetWordVersion();
	BOOL IsColor(COLORREF color);
	BOOL m_bQuit;//是否负责Word退出
	void Set_ShowParagraphs(BOOL bShow=TRUE);
	void FunAreaMin();
	void CloseWordToolBar(BOOL bAll);
	int GetDocument_CountLines(int ParamCode=23);
	CString  CyzWordOperator::GetDocument_CountInfo();

	void GetWholeStory(int& iStart,int& iEnd);
	BOOL UserFind(CString szKeyword,int& iStart);
	void SetGraphCenter();
	void Replace_UnderLine(CString FindStr,CString ReplStr);
	BOOL Select_Seg(int iStart, int iEnd);
	void FindKeywordParam(CString szKeyword,CDWordArray& StartAry);
	void DeletePageHeader();
	void SetDocSave(BOOL bSave=TRUE);
	BOOL FindReplace(CString Keyword, char Space);
	BOOL FindReplace(CString Keyword,CString replace,short reps=2);
	BOOL Replace_Space_Tab();
	void SetTabPos(float pos);
	void SetParagraphFormat();
	CString GetStartAndEndText(int iStart,int iEnd);
	void SearchKeyword(CString &keyword, int &iStart, int iEnd);
	BOOL ResorteTxt(int iStart, int iEnd);
	void FooteSwitch();
	BOOL SetSelectTxtColor(DWORD clr,BOOL bBOld=TRUE,int chars=1);
	BOOL SetSelectTxtColor_New(DWORD clr,int Pos,BOOL bBOld=TRUE,int chars=1);
	BOOL IsMainPane(int& nPanes);
	BOOL UpProtectedDoc(CString password);
	CString GetCurrentLineText(int& pos,int &CurPos);
	void SetStartEnd(int iStart, int iEnd);
	BOOL GetStartEndPos(int &iStart, int &iEnd);
	BOOL SearchKeywordPos(CString &keyword);
	BOOL ProtectedDoc(CString password);
	void DeleteAll(void);
	void SetTrimSelect(int iStart, int iEnd);
	void SetSelectPos(int  iStart, int  iEnd);
	void SearchKeywordStartPos(CString  szKeyword, CDWordArray& StartAry);
	BOOL SearchKeyword(CString& keyword);

	void FreeWordObj();
	void ReleaseDispatch();
	BOOL GetDocSaved();
	IDispatch* m_pdispWordApp;

protected:
	BOOL InitOLE();
	BOOL AddArgumentBool(LPOLESTR lpszArgName, WORD wFlags, BOOL b);
	BOOL SetWordVisible(BOOL bVisible);
	BOOL CreateBlankDocument();
	void ShowException(LPOLESTR szMember, HRESULT hr, EXCEPINFO *pexcep, unsigned int uiArgErr);
	int	m_iArgCount;
	int	m_iNamedArgCount;
	VARIANTARG	m_aVargs[MAX_DISP_ARGS];
	DISPID		m_aDispIds[MAX_DISP_ARGS + 1];		// one extra for the member name
	LPOLESTR	m_alpszArgNames[MAX_DISP_ARGS + 1];	// used to hold the argnames for GetIDs

	BOOL WordInvoke(IDispatch *pdisp, LPOLESTR szMember, VARIANTARG * pvargReturn,
			WORD wInvokeAction, WORD wFlags);
public:
	bool Wordexit();
};
class CyzWordTable : public CyzWordOperator  
{
public:
	void SetRowFormat(WORD_NAMESPACE::Table &tab,int irow);
	long GetEndPos(WORD_NAMESPACE::Table& table);
	void SetColumnForat(WORD_NAMESPACE::Table tab,int Col,int nFormat=1);
	void SetColumnWide(WORD_NAMESPACE::Table& table,int Col,float wide);
	void PasteCells(WORD_NAMESPACE::Table& tab,CPoint start,CPoint end);
	void SetBorders(WORD_NAMESPACE::Table TabObj,CPoint pt);
	void SetCellFormat(WORD_NAMESPACE::Table tab,CPoint cell,int nFormat=1);
	void SetCellText(WORD_NAMESPACE::Table tab,CPoint x,CString txt);
	BOOL HebinCell(WORD_NAMESPACE::Table TableObj,CPoint start,CPoint end);
	WORD_NAMESPACE::Table AddTable(int nRows,int nCols);
	CyzWordTable();
	virtual ~CyzWordTable();
	WORD_NAMESPACE::Table m_CurTable;


};
class CWordHebin : public CObject  
{
public:
	CWordHebin();
	virtual ~CWordHebin();
	CMap<WORD,WORD,_Document*,_Document*>m_NoDocment;
public:
	void SetMargin();
	void InsertTable(CyzWordTable& Table,int Rows,int Cols,CPtrArray& wides);
	void SetParagraphFormat();
	void MoveEnd();
	void SetVisible(BOOL bVisible);
	BOOL InsertBreakPage();
	void OnOpenDoc_Copy(CString szDocFileName);
	BOOL Insert_DocFile(CString szDocName,int nDocNo=1);
	void Insert_TypeText(CString txt,CString szFontName,int Alignment,float size,int nDocNo=1,DWORD Color=0,BOOL bBold=FALSE);
	void SaveDocument(CString DocName,int nDocNo=1);
	BOOL CreateWordApp(int nDocs=1);
	WORD_NAMESPACE::_Application* m_TargetWord_App;//目标Word应用程序对象
protected:
	BOOL m_nCloseApp;//是否关闭Word应用程序
	WORD_NAMESPACE::Selection* m_Cur_Sel;//当前选定位置
	int m_nDocs;//创建的文档数
	//	int m_nDoc;//当前文档序号
	static _Document CurDoc;//=NULL;
	static _Document CurDaAn_Doc;//=NULL;//答案
	static _Document CurDaAn_Txt;//=NULL;//答案
	static _Document CurSTDA_Doc;//=NULL;//试案
	

};

#endif // !defined(AFX_YZWORDOPERATOR_H__C4696929_36D6_4DCC_9A81_91DBCF67A837__INCLUDED_)
