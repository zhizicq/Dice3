// yzWordOperator.cpp: implementation of the CyzWordOperator class.
//
//////////////////////////////////////////////////////////////////////
#include "pch.h"
#include "stdafx.h"
#include "yzWordOperator.h"
#include "comdef.h"
#include <ole2ver.h>
#include "WordOperator.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////
using namespace WORD_NAMESPACE;

CyzWordOperator::CyzWordOperator()
{
	m_pWordApp=NULL;
	m_WordDoc=NULL;
	m_pStParamTab=NULL;
	m_bQuit=FALSE;
	m_pdispWordApp=NULL;
	m_iArgCount=0;
	m_MaxPos=-1;
		
	//CoInitialize(NULL) != S_OK
	//CoUninitialize();
}

CyzWordOperator::~CyzWordOperator()
{
	if(m_bQuit&&m_pWordApp)
	{
		//退出
		m_pWordApp->m_bAutoRelease=TRUE;
		VARIANT vt ;
		vt.vt =VT_ERROR;
		vt.scode =DISP_E_PARAMNOTFOUND;
		
		VARIANT v;
		v.vt =VT_BOOL;
		v.boolVal =VARIANT_FALSE;
		m_pWordApp->Quit(&v,&vt,&vt);
		m_pWordApp->DetachDispatch();
		m_pWordApp->ReleaseDispatch();
		
		//COleVariant vtMissing(DISP_E_PARAMNOTFOUND, VT_ERROR); 
		//BYTE parms[] =VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
		//m_pWordApp->InvokeHelper(0x451, DISPATCH_METHOD, VT_EMPTY, NULL, parms,&vtMissing, &vtMissing, &vtMissing); 

		delete m_pWordApp;
	}
	else if(m_pWordApp)
	{
		m_pWordApp->ReleaseDispatch();
		if(m_WordDoc)
			m_WordDoc.ReleaseDispatch();
		delete m_pWordApp;
		m_pWordApp=NULL;
	}
}
BOOL CyzWordOperator::savePicture(HWND hWnd, CString FileName)
{
	OpenClipboard(hWnd);

	if (IsClipboardFormatAvailable(CF_BITMAP))
	{
		HANDLE hbitmap;
		HBITMAP bitmap;
		CImage nImage;
		hbitmap = GetClipboardData(CF_BITMAP);
		bitmap = (HBITMAP)hbitmap;
		//				CBitmap nbmp;	
		nImage.Attach(bitmap);
		nImage.Save(FileName + ".jpg");

	}

	CloseClipboard();
	return 0;
}


int CyzWordOperator::GetMaths(HWND hWnd,OMaths maths ,InlineShapes isps ) {
	long i = 0;
	if (m_pWordApp == NULL)
		return -1;

	if (isps == NULL||isps.m_lpDispatch == NULL)
		;
	else {
		long count = isps.GetCount();
		if (!count);
		else {
			
			for (long g = 0; g < count; g++)
			{
				CString FileName;
				FileName.Format("math_%d", i + 1);


				InlineShape ishape = isps.Item(g + 1);
				long type = ishape.GetType();

				if (ishape.GetOLEFormat() != NULL)
				{
					OLEFormat ole = ishape.GetOLEFormat();
					CString  ls = ole.GetClassType();
					int go = ls.Find("Equation");
					if (ls.Find("Equation.3")|| ls.Find("Equation.DSMT4"))
					{
						Selection m_Target_Sel = m_pWordApp->GetSelection();
						ishape.Select();
						// 复制到剪切板
						m_Target_Sel.CopyAsPicture();
						i++;
						savePicture(hWnd, FileName);
					}
				}

			}
		}
	}
	Selection m_Target_Sel = m_pWordApp->GetSelection();
	long count = maths.get_Count();
	if (maths == NULL)return i;
	if (!count)return i;
	for (long g = 1; g <= count; g++)
	{

		OMath math = maths.Item(g);

		Range ran = math.get_Range();
		ran.Select();

		ran.CopyAsPicture();
		CString FileName;
		FileName.Format("math_%d", i);
		i++;
		OpenClipboard(hWnd);

		static int cfid = 0;
		cfid = RegisterClipboardFormat("HTML Format");
		if (IsClipboardFormatAvailable(cfid))
		{
			HANDLE hbitmap;

			hbitmap = GetClipboardData(cfid);
			char* p = (char*)GlobalLock(hbitmap);
			CString str = p;
			GlobalUnlock(hbitmap);
			CloseClipboard();
			int start = str.Find("src");
			int end = str.Find("png");
			CString path = str.Mid(start + 13, end - start - 10);
			CImage nImage;
			//				CBitmap nbmp;	
			nImage.Load(path);
			nImage.Save(FileName + ".jpg");
		}
		CloseClipboard();
	}
	return i;
}
BOOL CyzWordOperator::GetWordAppObj()
{
	BOOL bSuccess=FALSE;
	CLSID clsid;
	//获取Word服务程序的类标识
	if(SUCCEEDED(CLSIDFromProgID(OLESTR("Word.Application"),&clsid)))
	{
		IUnknown* pUnknown;
		//查询系统是否有Word应用程序是否存在
		if(SUCCEEDED(::GetActiveObject(clsid,NULL,&pUnknown)))
		{
			IDispatch* pDispatch;
			//查询Word应用程序接口
			if(SUCCEEDED(pUnknown->QueryInterface(IID_IDispatch,(void**)&pDispatch)))
			{
				pDispatch->Release();
				//建立关联的_Application对象
				if(!m_pWordApp)
					m_pWordApp=new _Application;
				else
					m_pWordApp->ReleaseDispatch();
				//引用Word应用程序
				m_pWordApp->AttachDispatch(pDispatch);
				//获取激活文档
				TRY
				{
					m_WordDoc=m_pWordApp->GetActiveDocument();
				    m_WordCom = m_pWordApp->GetCommandBars();
				}
				CATCH_ALL(e)
				{
					m_WordDoc=NULL;//没有文档
				}
				END_CATCH_ALL
				return TRUE;
			}
		}
	}
	return FALSE;
}
Shape CyzWordOperator::GetWordShapes()
{
	ShapeRange shaper = NULL;
	if (m_pWordApp == NULL)
		return NULL;
	Shapes shapes = m_WordDoc.GetShapes();
	InlineShapes Inlshapes = m_WordDoc.GetInlineShapes();
	if (shapes.GetCount() == 0)
		return NULL;
	return shapes;

}
void CyzWordOperator::ReportInfo()
{
	if(m_pWordApp)
		AfxMessageBox("找到当前运行的Word对象了!");
}
//获取激活Word文档的行数
int CyzWordOperator::GetWordDocLines(BOOL bLine)
{
	int nLines=-1;
	if(m_pWordApp==NULL)
		return -1;
	Documents myDocs; 
	_Document myDoc; 
	myDocs=m_pWordApp->GetDocuments();
	myDoc=m_pWordApp->GetActiveDocument();
	//以下一段代码取出Word文档的行数或页数
	CString   sProperty("Number of lines");//"Number of pages");
	if(!bLine)
		sProperty="Number of pages";
	LPDISPATCH   lpdispProps;   
	lpdispProps   =   myDoc.GetBuiltInDocumentProperties();   
    
	//Get   the   requested   Item   from   the   BuiltinDocumentProperties     
	//collection   
	//NOTE:     The   DISPID   of   the   "Item"   property   of   a     
	//               DocumentProperties   object   is   0x0   
	VARIANT   vResult;   
	DISPPARAMS   dpItem;   
	VARIANT   vArgs[1];   
	vArgs[0].vt   =   VT_BSTR;   
	vArgs[0].bstrVal   =   sProperty.AllocSysString();   
	dpItem.cArgs=1;   
	dpItem.cNamedArgs=0;   
	dpItem.rgvarg   =   vArgs;   
	HRESULT   hr   =   lpdispProps->Invoke(0x0,   IID_NULL,     
		LOCALE_USER_DEFAULT,   DISPATCH_PROPERTYGET,     
		&dpItem,   &vResult,   NULL,   NULL);   
	::SysFreeString(vArgs[0].bstrVal);   
    //Get   the   Value   property   of   the   BuiltinDocumentProperty   
	//NOTE:     The   DISPID   of   the   "Value"   property   of   a     
	//DocumentProperty   object   is   0x0   
	DISPPARAMS   dpNoArgs   =   {NULL,   NULL,   0,   0};   
	LPDISPATCH   lpdispProp;   
	lpdispProp   =   vResult.pdispVal;   
	hr   =   lpdispProp->Invoke(0x0,   IID_NULL,   LOCALE_USER_DEFAULT,     
		DISPATCH_PROPERTYGET,   &dpNoArgs,   &vResult,     
		NULL,   NULL);   
    
	//Set   the   text   in   the   Edit   Box   to   the   property's   value   
	CString   sPropValue   =   "";   
	switch   (vResult.vt)   
	{   
	case   VT_BSTR:   
		sPropValue   =   vResult.bstrVal;   
		break;   
	case   VT_I4:
		nLines=vResult.lVal;
		sPropValue.Format("%d",vResult.lVal);   
		break;   
	case   VT_DATE:   
		{   
			COleDateTime   dt   (vResult);   
			sPropValue   =   dt.Format(0,   LANG_USER_DEFAULT);   
			break;   
		}   
	default:   
		sPropValue   =   "<Information   for   the   property   you   selected    is   not   available>";   
		break;   
	}
	return nLines;
	
	
}

void CyzWordOperator::SelectRange(int nStartLine, int nEndLine)
{
	if(m_pWordApp==NULL)
		return ;
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	COleVariant What((short)3),Which((short)1),Continue((short)nStartLine),Name("");
	//插入光标移动到指定行
	COleVariant Unit((short)5),Extend((short)TRUE);
	m_Target_Sel.GoTo(&What,&Which,&Continue,&Name);
	m_Target_Sel.HomeKey(&Unit,&Extend);
	Continue.intVal=nEndLine-nStartLine;
	m_Target_Sel.MoveDown(&Unit,&Continue,&Extend);
	m_Target_Sel.EndKey(&Unit,&Extend);//移动到行尾，Unit=wdLine=5,Exten=TRUE选中
	m_Target_Sel.Copy();
//	CString txt=m_Target_Sel.GetText();
//	AfxMessageBox(txt);
	Unit.intVal=4;//段标识
	m_Target_Sel.StartOf(&Unit,&Extend);//移到段头
	m_Target_Sel.EndOf(&Unit,&Extend);//移到段头
	Continue.intVal=1;
	long k=m_Target_Sel.MoveDown(&Unit,&Continue,&Extend);

	
}

void CyzWordOperator::SetZoom(long bl)
{
	if(m_pWordApp==NULL)
		return ;

	//改变视图显示比例
	//改变视图显示比例
	Window WordWin;
	Pane pane1;
	WordWin.AttachDispatch(m_pWordApp->GetActiveWindow());
	pane1.AttachDispatch(WordWin.GetActivePane());
	View view;
	view.AttachDispatch(pane1.GetView());
	Zoom zoom;
	zoom.AttachDispatch(view.GetZoom());
	if(bl<100)
		zoom.SetPageFit(2);
	else
		zoom.SetPercentage(bl);
	if(WordWin.GetDisplayRulers())
		WordWin.SetDisplayRulers(0);

/*
HWND   hwnd   =     FindWindowEx(m_hWnd,NULL,"EXCEL2",NULL);   
  HWND   hclosewnd   =   NULL;   
    
  while   (hwnd   !=   NULL)   
  {   
  hclosewnd   =   FindWindowEx(hwnd,NULL,"MsoCommandBar","工作表菜单栏");   
  if   (hclosewnd)   
  SendMessage(hclosewnd,WM_CLOSE,0,0);   
  hclosewnd   =   FindWindowEx(hwnd,NULL,"MsoCommandBar","图表菜单栏");   
  if   (hclosewnd)   
  SendMessage(hclosewnd,WM_CLOSE,0,0);   
  hclosewnd   =   FindWindowEx(hwnd,NULL,"MsoCommandBar","格式");   
  if   (hclosewnd)   
  SendMessage(hclosewnd,WM_CLOSE,0,0);   
  hclosewnd   =   FindWindowEx(hwnd,NULL,"MsoCommandBar","常用");   
  if   (hclosewnd)   
  SendMessage(hclosewnd,WM_CLOSE,0,0);   
  hclosewnd   =   FindWindowEx(hwnd,NULL,"MsoCommandBar","图表");   
  if   (hclosewnd)   
  SendMessage(hclosewnd,WM_CLOSE,0,0);   
*/
}

void CyzWordOperator::SetSelectRangFont()
{
	if(m_pWordApp==NULL)
		return ;

	Selection m_Target_Sel=m_pWordApp->GetSelection();//.GetSelection();
	//m_Target_Sel.WholeStory();
	_Font m_wdFt = m_Target_Sel.GetFont();
	//m_Target_Sel.SetText("F");

	m_wdFt.SetSize(10.5);
	m_wdFt.SetName("宋体");
	m_wdFt.SetNameAscii("Times New Roman");
	m_wdFt.SetNameOther("Times New Roman");
	m_Target_Sel.SetFont(m_wdFt.DetachDispatch());

}
//确定参数行位置
BOOL CyzWordOperator::FindKeyWord(CString &keyword)
{
	if(m_pWordApp==NULL)
		return FALSE;
	m_StInfoInDoc.FreeMapTab();
	//取出总行数
	int nmax=GetWordDocLines();
	m_StInfoInDoc.m_Doc_Max_Line=nmax;
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	COleVariant What((short)3),Which((short)1),Continue((short)1),Name("");
	//插入光标移动到指定行
	COleVariant Unit((short)5),Extend((short)TRUE);
	COleVariant	Duan_Flag((short)4);//段标识
	for(int nL=1;nL<=nmax;nL++)
	{
		Continue.intVal=nL;
		m_Target_Sel.GoTo(&What,&Which,&Continue,&Name);
		m_Target_Sel.EndOf(&Unit,&Extend);//移到段尾;
		CString str=m_Target_Sel.GetText();
		//读取试题编号
		if(str.Find(keyword,0)>=0)
		{
			int i;
			CString str1,str2,str3,str4;
			int idx=str.Find(keyword);
			for(i=idx;i>=0&&i<str.GetLength();i++)
			{
				char chr=str.GetAt(i);
				if((chr>='0'&&chr<='9')||(chr>='a'&&chr<='z')||(chr>='A'&&chr<='Z'))
					str1+=chr;
				if(str1.GetLength()>=6)
					break;
			}
			//取分值
			idx=str.Find(str1);//找出试题编号存在的位置
			str2=str.Mid(idx+str1.GetLength());
			str3.Empty();
			for(i=0;i<str2.GetLength();i++)
			{
				char chr=str2.GetAt(i);
				if((chr>='0'&&chr<='9')||chr=='.')
					str3+=chr;
			}


			m_Target_Sel.GoTo(&What,&Which,&Continue,&Name);
			COleVariant Page=m_Target_Sel.GetInformation(1);//取当前页号//(10);
			COleVariant Line=m_Target_Sel.GetInformation(10);//取当前行号//(10);
			str.Format("%s:%s",str1,str3);
			m_StInfoInDoc.AddParamLine(nL,Page.intVal,Line.intVal,str,m_pStParamTab);
		}
	}
	//确定知识点,题干和答案行
//	if(m_StInfoInDoc.m_DocSts.GetCount()>0)
//		Search_Keyword_et();
	return FALSE;
}
//在参数行间找出知识点存在的位置
void CyzWordOperator::Search_Keyword_et()
{
	int EndLine=0;
	int CurLine=0;
	CStWordSeg *pOperate_Obj=NULL;
	CStWordSeg *pWordSeg1=NULL;
	POSITION pos=m_StInfoInDoc.m_DocSts.GetHeadPosition();
	while(pos)
	{
		pWordSeg1=m_StInfoInDoc.m_DocSts.GetNext(pos);
		if(CurLine==0)
		{
			CurLine=pWordSeg1->m_Para_Line;
			EndLine=CurLine;
			pOperate_Obj=pWordSeg1;//当前处理行对象
			//找出下一个参数行号
			if(pos)
			{
				pWordSeg1=m_StInfoInDoc.m_DocSts.GetNext(pos);
				EndLine=pWordSeg1->m_Para_Line;
			}
			else
				break;
		}
		else
			EndLine=pWordSeg1->m_Para_Line;
		if(CurLine!=EndLine)
		{
			Find_KeyWord_Et(CurLine,EndLine,pOperate_Obj);
			CurLine=EndLine;
			pOperate_Obj=pWordSeg1;
		}
	}
	if(CurLine==EndLine&&pOperate_Obj)
	{
		EndLine=m_StInfoInDoc.m_Doc_Max_Line;
		Find_KeyWord_Et(CurLine,EndLine,pOperate_Obj);
	}
/*		
	pos=m_StInfoInDoc.m_DocSts.GetHeadPosition();
	while(pos)
	{
		CStWordSeg *pWordSeg=m_StInfoInDoc.m_DocSts.GetNext(pos);

		CString str;
		str.Format("参数行%d\n知识点:%d\n题干行:%d\n答案行:%d",pWordSeg->m_Para_Line,pWordSeg->m_keyword_Line,
			pWordSeg->m_Tigan_Line,pWordSeg->m_Daan_Line);
		AfxMessageBox(str);
	}
*/

}

BOOL CyzWordOperator::Find_KeyWord_Et(int StartLine, int EndLine, CStWordSeg *pObj)
{
	if(pObj==NULL)
		return FALSE;
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	COleVariant What((short)3),Which((short)1),Continue((short)1),Name("");
	//插入光标移动到指定行
	COleVariant Unit((short)5),Extend((short)TRUE);
	COleVariant	Duan_Flag((short)4);//段标识
	for(int nL=StartLine;nL<EndLine;nL++)
	{
		Continue.intVal=nL;
		m_Target_Sel.GoTo(&What,&Which,&Continue,&Name);
		m_Target_Sel.EndOf(&Unit,&Extend);//移到段尾;
		CString str=m_Target_Sel.GetText();
		if(str.Find("知识点",0)>=0)
			pObj->m_keyword_Line=nL;
		else if(str.Find("题干",0)>=0)
			pObj->m_Tigan_Line=nL;
		else if(str.Find("答案",0)>=0)
			pObj->m_Daan_Line=nL;
	}
	return TRUE;

}

void CyzWordOperator::DispResult(CListCtrl *pListCtrl)
{
	if(pListCtrl==NULL)
		return;
	pListCtrl->DeleteAllItems();
	int ns=pListCtrl->GetHeaderCtrl()->GetItemCount();
	int i;
	for(i=ns-1;i>=0;i--)
		pListCtrl->DeleteColumn(i);

	pListCtrl->InsertColumn(0,"序号",LVCFMT_LEFT,40);
	pListCtrl->InsertColumn(1,"题型编码",LVCFMT_LEFT,80);
	pListCtrl->InsertColumn(2,"难度编码",LVCFMT_LEFT,90);
	pListCtrl->InsertColumn(3,"分值",LVCFMT_LEFT,80);
	pListCtrl->InsertColumn(4,"章节编码",LVCFMT_LEFT,120);
	pListCtrl->InsertColumn(5,"试题参数行号",LVCFMT_LEFT,120);
	pListCtrl->InsertColumn(6,"试题知识点行号",LVCFMT_LEFT,120);
	pListCtrl->InsertColumn(7,"试题题干行号",LVCFMT_LEFT,120);
	pListCtrl->InsertColumn(8,"试题答案行号",LVCFMT_LEFT,120);

	
	
	
	POSITION pos=m_StInfoInDoc.m_DocSts.GetHeadPosition();
	i=0;
	int j=0;
	while(pos)
	{
		CStWordSeg *pWordSeg=m_StInfoInDoc.m_DocSts.GetNext(pos);
		CString szH,szKeyLine,szZsd,szTg,szDa;

		szH.Format("%d",i+1);
		pListCtrl->InsertItem(i,szH);

		pListCtrl->SetItemText(i,1,pWordSeg->m_Tx);
		pListCtrl->SetItemText(i,2,pWordSeg->m_Nd);
		szZsd.Format("%g",pWordSeg->m_Fz);
		pListCtrl->SetItemText(i,3,szZsd);
		pListCtrl->SetItemText(i,4,pWordSeg->m_Zj);
		j=5;

		szKeyLine.Format("%d",pWordSeg->m_Para_Line);
		szZsd.Format("%d",pWordSeg->m_keyword_Line);
		szTg.Format("%d",pWordSeg->m_Tigan_Line);
		szDa.Format("%d",pWordSeg->m_Daan_Line);
		pListCtrl->SetItemText(i,j++,szKeyLine);
		pListCtrl->SetItemText(i,j++,szZsd);
		pListCtrl->SetItemText(i,j++,szTg);
		pListCtrl->SetItemText(i,j++,szDa);
		i++;
	}


}

void CyzWordOperator::GotoLine(int keyline)
{
	if(m_pWordApp==NULL)
		return ;
	
	if(m_pWordApp==NULL)
		return ;
	
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	COleVariant What((short)3),Which((short)1),Continue((short)keyline),Name("");
	COleVariant Unit((short)5),Extend((short)TRUE);
	COleVariant	Duan_Flag((short)4);//段标识
	Continue.intVal=keyline;//关键字行
	COleVariant CharFlag((short)1);
	//插入光标移动到指定行
	m_Target_Sel.GoTo(&What,&Which,&Continue,&Name);

/*
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	COleVariant What((short)3),Which((short)1),Continue((short)1),Name("");
	COleVariant Unit((short)5),Extend((short)TRUE);
	COleVariant	Duan_Flag((short)4);//段标识
	Continue.intVal=keyline;//关键字行
	COleVariant CharFlag((short)1);
	//插入光标移动到指定行
	m_Target_Sel.GoTo(&What,&Which,&Continue,&Name);
	m_Target_Sel.ReleaseDispatch();
*/

	//取前行行号
/*	COleVariant Page=m_Target_Sel.GetInformation(1);//取当前页号//(10);
	COleVariant Line=m_Target_Sel.GetInformation(10);//取当前行号//(10);

	CString str;
	str.Format("%d页,%d行",Page.intVal,Line.intVal);
	AfxMessageBox(str);
*/
	return;
/*
	//行对应的CStWordSeg对象
	char szItemName[64];
	CStWordSeg* pSegObj=m_StInfoInDoc.FindKeyWordObj(keyline,szItemName);
	CStWordSeg* pNext=m_StInfoInDoc.FindNextObj(pSegObj);
	int nLine=keyline;
	int nEnd=m_StInfoInDoc.m_Doc_Max_Line;
	int Lines[5]={0};
	if(pSegObj)
	{
		Lines[0]=pSegObj->m_Para_Line;
		Lines[1]=pSegObj->m_keyword_Line;
		Lines[2]=pSegObj->m_Tigan_Line;
		Lines[3]=pSegObj->m_Daan_Line;
		if(pNext)
			Lines[4]=pNext->m_Para_Line;
		else
			Lines[4]=m_StInfoInDoc.m_Doc_Max_Line;
		//排序
		for(int i=0;i<4;i++)
		{
			int ls;
			for(int j=i+1;j<5;j++)
			{
				if(Lines[i]>Lines[j])
				{
					ls=Lines[i];
					Lines[i]=Lines[j];
					Lines[j]=ls;
				}
			}
		}
		//确定本次操作的起止行
		for(i=0;i<5;i++)
		{
			if(Lines[i]==keyline)
			{
				nEnd=Lines[i+1];
				break;
			}
		}
	}


	//取出总行数
	int nmax=GetWordDocLines();
	m_StInfoInDoc.m_Doc_Max_Line=nmax;
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	COleVariant What((short)3),Which((short)1),Continue((short)1),Name("");
	COleVariant Unit((short)5),Extend((short)TRUE);
	COleVariant	Duan_Flag((short)4);//段标识
	Continue.intVal=keyline;//关分键字行
	COleVariant CharFlag((short)1);
	//插入光标移动到指定行
	m_Target_Sel.GoTo(&What,&Which,&Continue,&Name);
	int nChars=strlen(szItemName);
	nChars=nChars/2+nChars%2;
	Extend.boolVal=FALSE;
	Continue.intVal=nChars;
	m_Target_Sel.MoveRight(&CharFlag,&Continue,&Extend);
	CString txt=m_Target_Sel.GetText();
	CString szbiao="!:,.：、。，！";
	if(szbiao.Find(txt,0)>=0)
	{
		Continue.intVal=1;
		m_Target_Sel.MoveRight(&CharFlag,&Continue,&Extend);
	}

	Extend.boolVal=TRUE;

	//计算机连续移动行数
	Continue.intVal=nEnd-keyline-1;
	m_Target_Sel.MoveDown(&Unit,&Continue,&Extend);
	m_Target_Sel.EndOf(&Duan_Flag,&Extend);//移到段尾;

	//复制到Word粘贴板中
	m_Target_Sel.Copy();
*/
}
//取当前位置处的一个字
CString CyzWordOperator::GetText()
{
	CString txt("");
	if(m_pWordApp==NULL)
		return txt;
	
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	txt=m_Target_Sel.GetText();
	return txt;
}
//取当前光标处的单词
CString CyzWordOperator::GetWords(int &chars)
{
	CString txt("");
	chars=0;
	if(m_pWordApp==NULL)
		return txt;
	try
	{
		Selection m_Target_Sel=m_pWordApp->GetSelection();
		int start=m_Target_Sel.GetStart();
		int end=m_Target_Sel.GetEnd();
		Words words=m_Target_Sel.GetWords();//.GetText();取单词数
		Selection sel=words.Item(1);//第一个单词
		txt=sel.GetText();
		Characters charObj=sel.GetCharacters();
		chars=charObj.GetCount();
		m_Target_Sel.SetStart(start);
		m_Target_Sel.SetEnd(end);
		for(int i=1;i<chars;i++)
		{
			Range Range=words.Item(i);
			_Font Font=Range.GetFont();
			//CString cstext=pRange->GetText();
			long crColor=Font.GetColor();//获取字体颜色，正常是正数，可是当在Word中设置字体为主题颜色的时候，获取值为负值，不是正常的颜色值。
			long p=Font.GetColorIndex();
		}

/*
		_Font font=m_Target_Sel.GetFont();
		if(font.m_lpDispatch!=NULL)
		{
			float size=font.GetSize();
			long color=font.GetColor();
		}
*/


		return txt;
	}
	catch(CException* e)
	{
		e->Delete();
	}
	return txt;
}

//取当前页号和行号
int CyzWordOperator::GetPageLine(int& nLine)
{
	if(m_pWordApp==NULL)
		return -1;
	try
	{
		Selection m_Target_Sel=m_pWordApp->GetSelection();

		//取前行行号
		COleVariant Page=m_Target_Sel.GetInformation(1);//取当前页号
		COleVariant Line=m_Target_Sel.GetInformation(10);//取当前行号
		nLine=Line.intVal;
		return Page.intVal;
	}
	catch(CException* e)
	{
		e->Delete();
		return -1;
	}
}
//
void CyzWordOperator::DispStParam(CListCtrl *pListCtrl)
{
	if(pListCtrl==NULL)
		return;
	pListCtrl->DeleteAllItems();
	int i=0, j=0;
	POSITION pos=m_pStParamTab->GetHeadPosition();
	while(pos)
	{
		CStWordSeg *pWordSeg=m_pStParamTab->GetNext(pos);
		CString szH,szZsd;

		szH.Format("%d",i+1);
		pListCtrl->InsertItem(i,szH);

		pListCtrl->SetItemText(i,1,pWordSeg->m_Tx);
		szZsd.Format("%g",pWordSeg->m_Fz);
		pListCtrl->SetItemText(i,2,szZsd);
		//pListCtrl->SetItemData(i,(DWORD)pWordSeg);
		i++;
	}
}

BOOL CyzWordOperator::ProtectedDoc()
{
	if(m_pWordApp==NULL)
		return FALSE;

	TRY
	{
		_Document myDoc; 
		myDoc=m_pWordApp->GetActiveDocument();
		Selection m_Target_Sel=m_pWordApp->GetSelection();
		Editors editors=m_Target_Sel.GetEditors();
		COleVariant EditorID((short)-1);
		editors.Add(&EditorID);
		
		COleVariant NoReset((short)false),UseIRM((short)false),EnforceStyleLock((short)false);
		COleVariant Password("chen");
		
		//	Protect(long Type, VARIANT* NoReset, VARIANT* Password, VARIANT* UseIRM, VARIANT* EnforceStyleLock)
		myDoc.Protect(3,&NoReset,&Password,&UseIRM,&EnforceStyleLock);
		SetDisplayPageBoundaries();
	}
	CATCH(CException, e)
	{
		return FALSE;
	}
	END_CATCH
	return TRUE;

}
//是否显示页边距
void CyzWordOperator::SetDisplayPageBoundaries(BOOL bVisable)
{
	if(m_pWordApp==NULL)
		return ;
	Window ActiveWindow=m_pWordApp->GetActiveWindow();
	View Viewa=ActiveWindow.GetView();
	Viewa.SetDisplayPageBoundaries(0);

}

void CyzWordOperator::SelectedCopy()
{
	if(m_pWordApp==NULL)
		return ;
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	m_Target_Sel.Copy();
}

void CyzWordOperator::Insert_TypeText(CString txt,CString szFontName,int Alignment,float size)
{
	if(m_pWordApp==NULL)
		return ;
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	_Font OldFont=m_Target_Sel.GetFont();
	if(size>0.0)
	{
		OldFont.SetSize(size);
		OldFont.SetName(szFontName);
		OldFont.SetNameAscii("Times New Roman");
	}
	_ParagraphFormat ParagFmt=m_Target_Sel.GetParagraphFormat();
	ParagFmt.SetAlignment(Alignment);//居中
	m_Target_Sel.TypeText(txt);
	m_Target_Sel=m_pWordApp->GetSelection();
	ParagFmt.SetAlignment(3);//左对齐
}
void CyzWordOperator::Insert_TypeText(CString txt, CString szFontName, int Alignment, float size,DWORD Color,BOOL bBold)
{
	if(m_pWordApp==NULL)
		return ;

	Selection Sel=m_pWordApp->GetSelection();
	_Font OldFont=Sel.GetFont();
	if(size>0.0)
	{
		if(bBold)
			OldFont.SetBold(1);
		OldFont.SetSize(size);
		OldFont.SetName(szFontName);
		OldFont.SetNameAscii("Times New Roman");
		OldFont.SetColor(Color);
	}
	_ParagraphFormat ParagFmt=Sel.GetParagraphFormat();
	ParagFmt.SetAlignment(Alignment);//1-居中
	Sel.TypeText(txt);
//	ParagFmt.SetAlignment(3);//3-左对齐
}
void CyzWordOperator::Insert_DocFile(CString szDocName)
{
	if(m_pWordApp==NULL)
		return ;
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	COleVariant vFalse((long)0),vTrue((long)1);
	COleVariant vNull("");
	m_Target_Sel.InsertFile(szDocName,&vNull,vFalse,vFalse,vFalse);
	//COleVariant wdCharacter((short)1), Count((short)(1));
	//m_Target_Sel.Delete(wdCharacter,Count);
	GotoLine(1);//插入光标移至第1行
	m_Target_Sel.ReleaseDispatch();
	
}

void CyzWordOperator::TypeTextCenter(CString txt,float size)
{
	if(m_pWordApp==NULL)
		return ;
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	_ParagraphFormat ParagFmt=m_Target_Sel.GetParagraphFormat();
	ParagFmt.SetAlignment(1);//居中
	_Font OldFont=m_Target_Sel.GetFont();
	if(size>0.0)
	{
		OldFont.SetSize(size);
		OldFont.SetName("华文楷体");
		OldFont.SetNameAscii("Times New Roman");
	}
	m_Target_Sel.TypeText(txt);
	m_Target_Sel=m_pWordApp->GetSelection();
	ParagFmt.SetAlignment(3);//左对齐


}
//_Application
void CyzWordOperator::WriteText(CString txt, WORD_NAMESPACE::_Application *pApp)
{
	if(pApp)
	{
		Documents Docs=pApp->GetDocuments();
		int ns=Docs.GetCount();
		if(ns==0)
		{
			COleVariant vFalse((long)0),vTrue((long)1);	
			COleVariant Template("");//E:\\ENNORMAL.DOT
			COleVariant NewTemplate((short)FALSE),DocumentType((short)FALSE),Visible((short)TRUE);
			Docs.Add(&Template,&NewTemplate,&DocumentType, &Visible);
			
		}
		Selection m_Target_Sel=pApp->GetSelection();
		
			m_Target_Sel.TypeText(txt);

	}
}

void CyzWordOperator::All_Paste()
{
	if(m_pWordApp==NULL)
		return ;
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	m_Target_Sel.WholeStory();
	SetSelectRangFont();
	m_Target_Sel.Paste();
	GotoLine(1);
	//SetStartEnd(1,1);
	

}
void CyzWordOperator::Paste()
{
	if(m_pWordApp==NULL)
		return ;
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	try
	{
		m_Target_Sel.Paste();
	}
	catch(CException *e)
	{
		e->Delete();
	}
	if(::OpenClipboard(NULL))
	{
		::EmptyClipboard();
		::CloseClipboard();
	}
}
BOOL CyzWordOperator::StartWordApp()
{
	if(GetWordAppObj())
		return TRUE;
	m_pWordApp=new _Application;
	//创建应用对象(启动Word应用程序)
	if (!m_pWordApp->CreateDispatch("Word.Application", NULL))
	{
		AfxMessageBox("启动Word应用程序失败!\r\n机器上可能未安装Office系统程序!", MB_OK | MB_SETFOREGROUND); 
		return FALSE;
	}
	m_bQuit=TRUE;

	COleVariant vFalse((long)0),vTrue((long)1);	
	COleVariant Template("Normal");//E:\\ENNORMAL.DOT
	COleVariant NewTemplate((short)FALSE),DocumentType((short)0),Visible((short)TRUE);

	//m_pWordApp->SetVisible(TRUE);

	return TRUE;
}

CString CyzWordOperator::Get_All_Text()
{
	CString txt="";
	if(GetWordAppObj())
	{
		Selection nSel=m_pWordApp->GetSelection();
		Selection m_Target_Sel=m_pWordApp->GetSelection();
		m_Target_Sel.WholeStory();
		txt=m_Target_Sel.GetText();
		GotoLine(1);
	}
	return txt;
}
//设置Word的显示模式
void CyzWordOperator::SetPageView(int model)
{
	if(m_pWordApp==NULL)
		GetWordAppObj();
	if(m_pWordApp)
	{
		Window ActiveWindow=m_pWordApp->GetActiveWindow();
		View Viewa=ActiveWindow.GetView();
		Viewa.SetType(model);
	}

}
//插入分页符
BOOL CyzWordOperator::InsertBreakPage()
{
	if(m_pWordApp==NULL)
		return FALSE;
	TRY
	{
		Selection m_Target_Sel=m_pWordApp->GetSelection();
		COleVariant wdPageBreak((long)7);
		m_Target_Sel.InsertBreak(&wdPageBreak);
	}
	CATCH(CException, e)
	{
		return FALSE;
	}
	END_CATCH
		return TRUE;

}

BOOL CyzWordOperator::OpenWordFile(CString szFileName)
{
	//Documents.Open FileName:="C:\MyFiles\MyDoc.doc", ReadOnly:=True
	//if(NULL  == m_pdispWordApp)
	if (m_pWordApp == NULL && m_pdispWordApp != NULL)
		m_pdispWordApp = NULL;
	if (!StartWord())
	{
		
		return FALSE;
	}
	else
	{
		CString PathName=m_WordDoc.GetFullName();
		if(szFileName==PathName&& m_pWordApp)
		{
			m_WordDoc=m_pWordApp->GetActiveDocument();
			return TRUE;
		}
	}

	VARIANTARG varg1;
	ClearAllArgs();
	if (!WordInvoke(m_pdispWordApp, L"Documents", &varg1, DISPATCH_PROPERTYGET, 0))
		return FALSE;
	
	ClearAllArgs();
	AddArgumentCString(L"FileName", 0, szFileName);
	if (!WordInvoke(varg1.pdispVal, L"Open", NULL, DISPATCH_METHOD, 0))
		return FALSE;
	
	return TRUE;
}

void CyzWordOperator::ClearVariant(VARIANTARG *pvarg)
{
	pvarg->vt = VT_EMPTY;
	pvarg->wReserved1 = 0;
	pvarg->wReserved2 = 0;
	pvarg->wReserved3 = 0;
	pvarg->lVal = 0;

}

void CyzWordOperator::ClearAllArgs()
{
	int i;
	
	for (i = 0; i < m_iArgCount; i++) 
	{
		if (m_awFlags[i] & DISPARG_NOFREEVARIANT)
			// free the variant's contents based on type
			ClearVariant(&m_aVargs[i]);
		else
			ReleaseVariant(&m_aVargs[i]);
	}

	m_iArgCount = 0;
	m_iNamedArgCount = 0;
}

BOOL CyzWordOperator::AddArgumentCString(LPOLESTR lpszArgName, WORD wFlags, CString szStr)
{
	BSTR b;
	
	b = szStr.AllocSysString();
	if (!b)
		return FALSE;
	AddArgumentCommon(lpszArgName, wFlags, VT_BSTR);
	m_aVargs[m_iArgCount++].bstrVal = b;
	
	return TRUE;

}

void CyzWordOperator::AddArgumentCommon(LPOLESTR lpszArgName, WORD wFlags, VARTYPE vt)
{
	ClearVariant(&m_aVargs[m_iArgCount]);
	
	m_aVargs[m_iArgCount].vt = vt;
	m_awFlags[m_iArgCount] = wFlags;
	
	if (lpszArgName != NULL) 
	{
		m_alpszArgNames[m_iNamedArgCount + 1] = lpszArgName;
		m_iNamedArgCount++;
	}

}
BOOL CyzWordOperator::WordInvoke(IDispatch *pdisp, LPOLESTR szMember, VARIANTARG * pvargReturn,
			WORD wInvokeAction, WORD wFlags)
{
	HRESULT hr;
	DISPPARAMS dispparams;
	unsigned int uiArgErr;
	EXCEPINFO excep;

	// Get the IDs for the member and its arguments.  GetIDsOfNames expects the
	// member name as the first name, followed by argument names (if any).
	m_alpszArgNames[0] = szMember;
	hr = pdisp->GetIDsOfNames( IID_NULL, m_alpszArgNames,
								1 + m_iNamedArgCount, LOCALE_SYSTEM_DEFAULT, m_aDispIds);
	if (FAILED(hr)) 
	{
		if (!(wFlags & DISP_NOSHOWEXCEPTIONS))
			ShowException(szMember, hr, NULL, 0);
		return FALSE;
	}
	
	if (pvargReturn != NULL)
		ClearVariant(pvargReturn);
	
	// if doing a property put(ref), we need to adjust the first argument to have a
	// named arg of DISPID_PROPERTYPUT.
	if (wInvokeAction & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) 
	{
		m_iNamedArgCount = 1;
		m_aDispIds[1] = DISPID_PROPERTYPUT;
		pvargReturn = NULL;
	}
	
	dispparams.rgdispidNamedArgs = m_aDispIds + 1;
	dispparams.rgvarg = m_aVargs;
	dispparams.cArgs = m_iArgCount;
	dispparams.cNamedArgs = m_iNamedArgCount;
	
	excep.pfnDeferredFillIn = NULL;
	
	hr = pdisp->Invoke(m_aDispIds[0], IID_NULL, LOCALE_SYSTEM_DEFAULT,
								wInvokeAction, &dispparams, pvargReturn, &excep, &uiArgErr);
	
	if (wFlags & DISP_FREEARGS)
		ClearAllArgs();
	
	if (FAILED(hr)) 
	{
		// display the exception information if appropriate:
		if (!(wFlags & DISP_NOSHOWEXCEPTIONS))
			ShowException(szMember, hr, &excep, uiArgErr);
	
		// free exception structure information
		SysFreeString(excep.bstrSource);
		SysFreeString(excep.bstrDescription);
		SysFreeString(excep.bstrHelpFile);
	
		return FALSE;
	}
	return TRUE;
}

void CyzWordOperator::ReleaseVariant(VARIANTARG *pvarg)
{
	VARTYPE vt;
	VARIANTARG *pvargArray;
	long lLBound, lUBound, l;
	
	vt = pvarg->vt & 0xfff;		// mask off flags
	
	// check if an array.  If so, free its contents, then the array itself.
	if (V_ISARRAY(pvarg)) 
	{
		// variant arrays are all this routine currently knows about.  Since a
		// variant can contain anything (even other arrays), call ourselves
		// recursively.
		if (vt == VT_VARIANT) 
		{
			SafeArrayGetLBound(pvarg->parray, 1, &lLBound);
			SafeArrayGetUBound(pvarg->parray, 1, &lUBound);
			
			if (lUBound > lLBound) 
			{
				lUBound -= lLBound;
				
				SafeArrayAccessData(pvarg->parray, (void**)&pvargArray);
				
				for (l = 0; l < lUBound; l++) 
				{
					ReleaseVariant(pvargArray);
					pvargArray++;
				}
				
				SafeArrayUnaccessData(pvarg->parray);
			}
		}
		else 
		{
			MessageBox(NULL, _T("ReleaseVariant: Array contains non-variant type"), "Failed", MB_OK | MB_ICONSTOP);
		}
		
		// Free the array itself.
		SafeArrayDestroy(pvarg->parray);
	}
	else 
	{
		switch (vt) 
		{
			case VT_DISPATCH:
				//(*(pvarg->pdispVal->lpVtbl->Release))(pvarg->pdispVal);
				pvarg->pdispVal->Release();
				break;
				
			case VT_BSTR:
				SysFreeString(pvarg->bstrVal);
				break;
				
			case VT_I2:
			case VT_BOOL:
			case VT_R8:
			case VT_ERROR:		// to avoid erroring on an error return from Excel
				// no work for these types
				break;
				
			default:
				MessageBox(NULL, _T("ReleaseVariant: Unknown type"), "Failed", MB_OK | MB_ICONSTOP);
				break;
		}
	}
	
	ClearVariant(pvarg);

}

void CyzWordOperator::ShowException(LPOLESTR szMember, HRESULT hr, EXCEPINFO *pexcep, unsigned int uiArgErr)
{
	TCHAR szBuf[512];
	
	switch (GetScode(hr)) 
	{
		case DISP_E_UNKNOWNNAME:
			wsprintf(szBuf, "%s: Unknown name or named argument.", szMember);
			break;
	
		case DISP_E_BADPARAMCOUNT:
			wsprintf(szBuf, "%s: Incorrect number of arguments.", szMember);
			break;
			
		case DISP_E_EXCEPTION:
			wsprintf(szBuf, "%s: Error %d: ", szMember, pexcep->wCode);
			if (pexcep->bstrDescription != NULL)
				lstrcat(szBuf, (char*)pexcep->bstrDescription);
			else
				lstrcat(szBuf, "<<No Description>>");
			break;
			
		case DISP_E_MEMBERNOTFOUND:
			wsprintf(szBuf, "%s: method or property not found.", szMember);
			break;
		
		case DISP_E_OVERFLOW:
			wsprintf(szBuf, "%s: Overflow while coercing argument values.", szMember);
			break;
		
		case DISP_E_NONAMEDARGS:
			wsprintf(szBuf, "%s: Object implementation does not support named arguments.",
						szMember);
		    break;
		    
		case DISP_E_UNKNOWNLCID:
			wsprintf(szBuf, "%s: The locale ID is unknown.", szMember);
			break;
		
		case DISP_E_PARAMNOTOPTIONAL:
			wsprintf(szBuf, "%s: Missing a required parameter.", szMember);
			break;
		
		case DISP_E_PARAMNOTFOUND:
			wsprintf(szBuf, "%s: Argument not found, argument %d.", szMember, uiArgErr);
			break;
			
		case DISP_E_TYPEMISMATCH:
			wsprintf(szBuf, "%s: Type mismatch, argument %d.", szMember, uiArgErr);
			break;

		default:
			wsprintf(szBuf, "%s: Unknown error occured.", szMember);
			break;
	}
	
	MessageBox(NULL, szBuf, "OLE Error", MB_OK | MB_ICONSTOP);
}

BOOL CyzWordOperator::CloseActiveDocument(BOOL bSaveIt)
{
	if(NULL  ==m_pWordApp->m_lpDispatch)
		return FALSE;
	int wdDoNotSaveChanges = 0;
	int wdPromptToSaveChanges = -2;
	int wdSaveOption = wdDoNotSaveChanges;
	if(bSaveIt)
		wdSaveOption = wdPromptToSaveChanges;

	VARIANTARG varg1;	
	ClearAllArgs();
	if (!WordInvoke(m_pWordApp->m_lpDispatch, L"ActiveDocument", &varg1, DISPATCH_PROPERTYGET, 0))
		return FALSE;
	ClearAllArgs();
	AddArgumentInt2(L"SaveChanges", 0, wdSaveOption);
	if (!WordInvoke(varg1.pdispVal, L"Close", NULL, DISPATCH_METHOD, 0))
		return FALSE;

	return TRUE;
}

BOOL CyzWordOperator::AddArgumentInt2(LPOLESTR lpszArgName, WORD wFlags, int i)
{
	AddArgumentCommon(lpszArgName, wFlags, VT_I2);
	m_aVargs[m_iArgCount++].iVal = i;
	return TRUE;

}

void CyzWordOperator::SetVisible(BOOL bVisible)
{
	if(m_pWordApp)
		m_pWordApp->SetVisible(bVisible);

}

BOOL CyzWordOperator::StartWord()
{
//	InitOLE();
	

	CLSID clsWordApp;


	// if Excel is already running, return with current instance
	if (m_pdispWordApp != NULL)
		return TRUE;


	/* Obtain the CLSID that identifies EXCEL.APPLICATION
	 * This value is universally unique to Excel versions 5 and up, and
	 * is used by OLE to identify which server to start.  We are obtaining
	 * the CLSID from the ProgID.
	 */
	if (FAILED(CLSIDFromProgID(L"Word.Application", &clsWordApp))) 
	{
		MessageBox(NULL, _T("不能获得Word的类标识符!"), "Failed", MB_OK | MB_ICONSTOP);
		return FALSE;
	}
	// start a new copy of Excel, grab the IDispatch interface
	if (FAILED(CoCreateInstance(clsWordApp, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&m_pdispWordApp))) 
	{
		MessageBox(NULL, _T("不能启动Word应用程序."), "Failed", MB_OK | MB_ICONSTOP);
		return FALSE;
	}
	SetWordVisible(TRUE);

	return TRUE;
}

BOOL CyzWordOperator::CreateBlankDocument()
{
	if(NULL  == m_pdispWordApp)
		return FALSE;

	VARIANTARG varg1, varg2;
	
	//Documents.Add
	ClearAllArgs();
	if (!WordInvoke(m_pdispWordApp, L"Documents", &varg1, DISPATCH_PROPERTYGET, 0))
		return FALSE;
	ClearAllArgs();
	if (!WordInvoke(varg1.pdispVal, L"Add", &varg2, DISPATCH_METHOD, 0))
		return FALSE;

	return TRUE;
}

BOOL CyzWordOperator::SetWordVisible(BOOL bVisible)
{
	if (m_pdispWordApp == NULL)
		return FALSE;
	
	ClearAllArgs();
	AddArgumentBool(NULL, 0, bVisible);
	return WordInvoke(m_pdispWordApp, L"Visible", NULL, DISPATCH_PROPERTYPUT, DISP_FREEARGS);

}

BOOL CyzWordOperator::AddArgumentBool(LPOLESTR lpszArgName, WORD wFlags, BOOL b)
{
	AddArgumentCommon(lpszArgName, wFlags, VT_BOOL);
	// Note the variant representation of True as -1
	m_aVargs[m_iArgCount++].boolVal = b ? -1 : 0;
	return TRUE;

}

BOOL CyzWordOperator::InitOLE()
{
	DWORD dwOleVer;
	
	dwOleVer = CoBuildVersion();
	
	// check the OLE library version
	if (rmm != HIWORD(dwOleVer)) 
	{
		MessageBox(NULL, _T("Incorrect version of OLE libraries."), "Failed", MB_OK | MB_ICONSTOP);
		return FALSE;
	}
	
	// could also check for minor version, but this application is
	// not sensitive to the minor version of OLE
	
	// initialize OLE, fail application if we can't get OLE to init.
	if (FAILED(OleInitialize(NULL))) 
	{
		MessageBox(NULL, _T("Cannot initialize OLE."), "Failed", MB_OK | MB_ICONSTOP);
		return FALSE;
	}
	
		
	return TRUE;
}



BOOL CyzWordOperator::GetDocSaved()
{
	if(m_pWordApp==NULL)
		return FALSE;
	_Document myDoc; 
	myDoc=m_pWordApp->GetActiveDocument();
	if(myDoc.GetSaved())
		return TRUE;
	return FALSE;
}

void CyzWordOperator::ReleaseDispatch()
{
	if(m_pWordApp)
	{
		m_pWordApp->ReleaseDispatch();
		delete m_pWordApp;
		m_pWordApp=NULL;
	}
}

void CyzWordOperator::FreeWordObj()
{
	if(m_bQuit&&m_pWordApp)
	{
		//退出
		m_pWordApp->m_bAutoRelease=TRUE;
		VARIANT vt ;
		vt.vt =VT_ERROR;
		vt.scode =DISP_E_PARAMNOTFOUND;
		
		VARIANT v;
		v.vt =VT_BOOL;
		v.boolVal =VARIANT_FALSE;
		m_pWordApp->Quit(&v,&vt,&vt);
		m_pWordApp->DetachDispatch();
		m_pWordApp->ReleaseDispatch();
		delete m_pWordApp;
	}
	else if(m_pWordApp)
	{
		m_pWordApp->ReleaseDispatch();
		delete m_pWordApp;
	}
	m_pWordApp=NULL;
}

void CyzWordOperator::SearchKeywordStartPos(CString szKeyword, CDWordArray &StartAry)
{
	StartAry.RemoveAll();//清空
	if(m_pWordApp==NULL)
		return;

		
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	m_Target_Sel.WholeStory();
	int endPos=m_Target_Sel.GetEnd();
	//从头开始
	//m_Target_Sel.SetStart(0);
	//m_Target_Sel.SetEnd(1);
	Find w_Find=m_Target_Sel.GetFind();
	w_Find.ClearFormatting();
	//m_Target_Sel.ClearFormatting();//清除查找对象
	COleVariant FindText(szKeyword);
	COleVariant MatchCase((short)0), MatchWholeWord((short)0), MatchWildcards((short)0);
	COleVariant MatchSoundsLike((short)0), MatchAllWordForms((short)0),Forward((short)1),Wrap((short)0);
	COleVariant Format((short)0),ReplaceWith(""), Replace((short)0), MatchKashida((short)0),MatchDiacritics((short)0);
	COleVariant MatchAlefHamza((short)0), MatchControl((short)0);

	int ns=0;
	while(w_Find.Execute(FindText,MatchCase, MatchWholeWord, MatchWildcards,MatchSoundsLike, 
		MatchAllWordForms,Forward,Wrap,Format,ReplaceWith, Replace, MatchKashida,MatchDiacritics,MatchAlefHamza,
		MatchControl))
	{
		DWORD nstart=m_Target_Sel.GetStart();
		StartAry.Add(nstart);
	}
	m_Target_Sel.SetStart(1);
	m_Target_Sel.SetEnd(1);
}

void CyzWordOperator::SetSelectPos(int iStart, int iEnd)
{
	if(m_pWordApp==NULL)
		return ;

	TRY
	{
		Selection m_Target_Sel=m_pWordApp->GetSelection();
/*		int nLine,nPage;
		
		m_Target_Sel.SetStart(iStart);

		nPage=GetPageLine(nLine);

		COleVariant What((short)1);
		COleVariant Which((short)1);
		COleVariant Count((short)nPage);
		COleVariant Name("");

		m_Target_Sel.GoTo(What,Which,Count,Name);
*/		
		m_Target_Sel.SetStart(iStart);
		m_Target_Sel.SetEnd(iEnd);
		m_Target_Sel.Select();

		//Window wdo=m_pWordApp->GetActiveWindow();
		//Pane pan=wdo.GetActivePane();
		//pan.AutoScroll(iStart);
		
		//m_Target_Sel.SelectRow();

	}
	CATCH(CException, e)
	{
		e->Delete();
		return ;
	}
	END_CATCH
}
BOOL CyzWordOperator::SearchKeyword(CString& keyword)
{
	if(m_pWordApp==NULL)
		return FALSE;

	try
	{
		Selection m_Target_Sel=m_pWordApp->GetSelection();
		//从头开始
		Find w_Find=m_Target_Sel.GetFind();
		//m_Target_Sel.ClearFormatting();//清除查找对象
		COleVariant FindText(keyword);
		COleVariant MatchCase((short)0), MatchWholeWord((short)0), MatchWildcards((short)0);
		COleVariant MatchSoundsLike((short)0), MatchAllWordForms((short)0),Forward((short)1),Wrap((short)1);
		COleVariant Format((short)0),ReplaceWith(""), Replace((short)0), MatchKashida((short)0),MatchDiacritics((short)0);
		COleVariant MatchAlefHamza((short)0), MatchControl((short)0);
		
		int ns=0;
		if(w_Find.Execute(FindText,MatchCase, MatchWholeWord, MatchWildcards,MatchSoundsLike, 
			MatchAllWordForms,Forward,Wrap,Format,ReplaceWith, Replace, MatchKashida,MatchDiacritics,MatchAlefHamza,
			MatchControl))
		{
			long lstart=m_Target_Sel.GetStart();
			m_Target_Sel.SetStart(lstart);
			m_Target_Sel.SetEnd(lstart);
			return TRUE;
		}
	}
	catch(CException* e)
	{
		e->Delete();
	}
	return FALSE;
}

void CyzWordOperator::SetTrimSelect(int iStart, int iEnd)
{
	if(m_pWordApp==NULL)
		return ;
	if(m_WordDoc)
		m_WordDoc.Activate();
	if(iStart>iEnd)
	{
		int ls=iStart;
		iStart=iEnd;
		iEnd=ls;
	}
	TRY
	{
		Selection m_Target_Sel=m_pWordApp->GetSelection();
		m_Target_Sel.SetStart(iStart);
		m_Target_Sel.SetEnd(iEnd);
		m_Target_Sel.Select();
		CString str=m_Target_Sel.GetText();
		if(str.GetLength()<=0)
			return;
		int iLen=str.GetLength();
		char* pStr=str.LockBuffer();
		char mh[]="：　";
		int i=0;
		int nChars=0;
		//过滤掉开始位置的换行符,全角冒号和全角空格
		do
		{
			if(pStr[i]==' '||pStr[i]=='\r'||pStr[i]=='\n'||pStr[i]==':')
			{
				i++;
				nChars++;
			}
			else if(((unsigned char)pStr[i])>0xA0)
			{
				if(pStr[i]==mh[0]&&pStr[i+1]==mh[1])
					i+=2,nChars++;
				else if(pStr[i]==mh[2]&&pStr[i+1]==mh[3])
					i+=2,nChars++;
				else
					break;
			}
			else
				break;

		}while(i<iLen);
		
		str.UnlockBuffer();

		//过滤掉后部无用字符
		int nEChars=0;
		int nEnter=0;
		int nEnter_f=0;
		str.Find("\r\n\r\n");
		i=iLen-1;

		while(i>0)
		{
			if(pStr[i]==' '||pStr[i]=='\r'||pStr[i]=='\n'||pStr[i]==':')
			{
				if(pStr[i]=='\r')
				{
					nEnter++;
					nEnter_f=nEChars;
				}
				i--;
				nEChars++;
			}
			else if(((unsigned char)pStr[i])>0xA0)
			{
				if(pStr[i]==mh[1]&&pStr[i-1]==mh[0])
					i-=2,nEChars++;
				else if(pStr[i]==mh[3]&&pStr[i-1]==mh[2])
					i-=2,nEChars++;
				else
					break;
			}
			else
				break;

		}
		if(nChars>0)
			iStart+=nChars;
		if(nEChars>0 && nEnter>0)
		{
			iEnd-=nEnter_f;
		}
		if(iStart<iEnd)
		{
			m_Target_Sel.SetStart(iStart);
			m_Target_Sel.SetEnd(iEnd);
			m_Target_Sel.Select();
		}

	}
	CATCH(CException, e)
	{
		e->Delete();
		return ;
	}
	END_CATCH
}
//删除文档数据
void CyzWordOperator::DeleteAll()
{
	if(m_pWordApp==NULL)
		return ;
	Documents Docs=m_pWordApp->GetDocuments();
	int iDocs=Docs.GetCount();
	if(Docs.GetCount()<=0)//无文档
		return;
	Docs.ReleaseDispatch();
	if(m_WordDoc!=NULL)//激活文档
		m_WordDoc.Activate();
	else
	{
		_Document Act_Doc=m_pWordApp->GetActiveDocument();
		Act_Doc.Activate();
		Act_Doc.ReleaseDispatch();
	}
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	m_Target_Sel.WholeStory();
	COleVariant wdCharacter((short)1), Count((short)1);//(end-start));
	m_Target_Sel.Delete(wdCharacter,Count);
	m_Target_Sel.WholeStory();
	m_Target_Sel.ReleaseDispatch();
	//DeletePageHeader();

}

BOOL CyzWordOperator::ProtectedDoc(CString password)
{
	if(m_pWordApp==NULL)
		return FALSE;

	TRY
	{
		_Document myDoc; 
		myDoc=m_pWordApp->GetActiveDocument();
		//Selection m_Target_Sel=m_pWordApp->GetSelection();
		////Editors editors=m_Target_Sel.GetEditors();
		////COleVariant EditorID((short)-1);
		//editors.Add(&EditorID);
		
		COleVariant NoReset((short)false),UseIRM((short)false),EnforceStyleLock((short)false),wdNoProtection((short)-1);
		COleVariant Password(password);
		
		//	Protect(long Type, VARIANT* NoReset, VARIANT* Password, VARIANT* UseIRM, VARIANT* EnforceStyleLock)
		if(myDoc.GetProtectionType()==-1)
			myDoc.Protect(3,&NoReset,&Password,&UseIRM,&EnforceStyleLock);
		//SetDisplayPageBoundaries();
	}
	CATCH(CException, e)
	{
		e->Delete();
		return FALSE;
	}
	END_CATCH
	return TRUE;
}

BOOL CyzWordOperator::SearchKeywordPos(CString &keyword)
{
	if(m_pWordApp==NULL)
		return FALSE;
	m_StInfoInDoc.FreeMapTab();
	//取出总行数
	int nmax=GetWordDocLines();
	m_StInfoInDoc.m_Doc_Max_Line=nmax;
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	COleVariant What((short)3),Which((short)1),Continue((short)1),Name("");
	//插入光标移动到指定行
	COleVariant Unit((short)5),Extend((short)TRUE);
	COleVariant	Duan_Flag((short)4);//段标识
	
	CDWordArray Starts;
	SearchKeywordStartPos(keyword,Starts);
	for(int is=0;is<Starts.GetSize();is++)
	{
		m_Target_Sel.SetStart(Starts.GetAt(is));
		m_Target_Sel.SetEnd(Starts.GetAt(is));
		m_Target_Sel.EndOf(&Unit,&Extend);//移到段尾;
		CString str=m_Target_Sel.GetText();
		//读取试题编号
		if(str.Find(keyword,0)>=0)
		{
			CString str1,str2,str3,str4;
			int idx=str.Find(keyword);
			int i;
			for(i=idx;i>=0&&i<str.GetLength();i++)
			{
				char chr=str.GetAt(i);
				if((chr>='0'&&chr<='9')||(chr>='a'&&chr<='z')||(chr>='A'&&chr<='Z'))
					str1+=chr;
				if(chr==']')
					break;
			}
			//取分值
			idx=str.Find(str1);//找出试题编号存在的位置
			str2=str.Mid(idx+str1.GetLength());
			str3.Empty();
			for(i=0;i<str2.GetLength();i++)
			{
				char chr=str2.GetAt(i);
				if((BYTE)chr>160)
					break;
				else if((chr>='0'&&chr<='9')||chr=='.')
					str3+=chr;
			}
			COleVariant Page=m_Target_Sel.GetInformation(1);//取当前页号//(10);
			COleVariant Line=m_Target_Sel.GetInformation(10);//取当前行号//(10);
			str.Format("%s:%s",str1,str3);//题号与分值
			//AfxMessageBox(str);
			//m_StInfoInDoc.AddParamLine(nL,Page.intVal,Line.intVal,str,m_pStParamTab);
			m_StInfoInDoc.AddParamStartPos(Starts.GetAt(is),Page.intVal,Line.intVal,str,m_pStParamTab);

		}
	}
	return FALSE;
}

BOOL CyzWordOperator::GetStartEndPos(int &iStart, int &iEnd)
{
	if(m_pWordApp==NULL)
		return FALSE;

	TRY
	{
		Selection m_Target_Sel=m_pWordApp->GetSelection();
		if(m_Target_Sel.m_lpDispatch!=NULL)
		{

			iStart=m_Target_Sel.GetStart();
			iEnd=m_Target_Sel.GetEnd();
			return TRUE;
		}
	
	}
	CATCH(CException, e)
	{
		e->Delete();
		return FALSE;
	}
	END_CATCH
	return TRUE;
}

void CyzWordOperator::SetStartEnd(int iStart, int iEnd)
{
	if(m_pWordApp==NULL||iStart<1||iEnd<1)
		return ;

	TRY
	{
		Selection m_Target_Sel=m_pWordApp->GetSelection();
		if(m_Target_Sel.m_lpDispatch!=NULL)
		{
			m_Target_Sel.SetStart(iStart);
			m_Target_Sel.SetEnd(iEnd);
		}
	
	}
	CATCH(CException, e)
	{
		e->Delete();
		return ;
	}
	END_CATCH
}
//取当前位置处的整行文本
CString CyzWordOperator::GetCurrentLineText(int &pos, int &CurPos)
{
	if(m_pWordApp==NULL)
		return "";
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	int nstart=m_Target_Sel.GetStart();
	int nend=m_Target_Sel.GetEnd();
	CurPos=nstart;
	CString txt;
	COleVariant Uint(short(5));//wdLine;
	COleVariant Extend(short(0));//wdExtend
	m_Target_Sel.HomeKey(&Uint,&Extend);
	int nhead=m_Target_Sel.GetStart();
	Extend.intVal=1;
	m_Target_Sel.EndKey(&Uint,&Extend);
	txt=m_Target_Sel.GetText();
	m_Target_Sel.SetStart(nstart);
	m_Target_Sel.SetEnd(nstart);
	pos=nstart-nhead;
	return txt;
}

BOOL CyzWordOperator::UpProtectedDoc(CString password)
{
	if(m_pWordApp==NULL)
		return FALSE;

	TRY
	{
		_Document myDoc; 
		myDoc=m_pWordApp->GetActiveDocument();
		///Selection m_Target_Sel=m_pWordApp->GetSelection();
		///Editors editors=m_Target_Sel.GetEditors();
		///COleVariant EditorID((short)-1);
		//editors.Add(&EditorID);
		
		///COleVariant NoReset((short)false),UseIRM((short)false),EnforceStyleLock((short)false),wdNoProtection((short)-1);
		COleVariant Password(password);
		
		//	Protect(long Type, VARIANT* NoReset, VARIANT* Password, VARIANT* UseIRM, VARIANT* EnforceStyleLock)
		if(myDoc.GetProtectionType()!=-1)
			myDoc.Unprotect(Password);//.Protect(3,&NoReset,&Password,&UseIRM,&EnforceStyleLock);
		//SetDisplayPageBoundaries();
	}
	CATCH(CException, e)
	{
		e->Delete();
		return FALSE;
	}
	END_CATCH
	return TRUE;
}

BOOL CyzWordOperator::IsMainPane(int &nPanes)
{
	if(m_pWordApp==NULL)
		return TRUE;
	TRY
	{
		Window win=m_pWordApp->GetActiveWindow();//激活的窗口
		Panes panes=win.GetPanes();
		nPanes=panes.GetCount();
		if(panes.GetCount()>1)
		{
			Pane ActivePane=win.GetActivePane();
			Pane pane2=panes.Item(1);//第1窗格为主窗格
			if(pane2!=ActivePane)
				return FALSE;
		}
	}
	CATCH(CException, e)
	{
		e->Delete();
		return TRUE;
	}
	END_CATCH
		return TRUE ;
}

BOOL CyzWordOperator::SetSelectTxtColor(DWORD clr, BOOL bBOld,int chars)
{
	if(m_pWordApp==NULL)
		return FALSE;
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	TRY
	{
		if(m_Target_Sel.m_lpDispatch!=NULL)
		{

			int ns=m_Target_Sel.GetStart();
			int ne=m_Target_Sel.GetEnd();

			if(ne>=ns && ns>0)
			{
				m_Target_Sel.SetStart(ns);
				m_Target_Sel.SetEnd(ns+chars);				
				_Font m_wdFt(m_Target_Sel.GetFont()); 
				m_wdFt.SetColor(clr);
				m_Target_Sel.SetFont(m_wdFt.DetachDispatch());
				//m_Target_Sel.SetStart(ns);
				//m_Target_Sel.SetEnd(ns);
				m_wdFt.ReleaseDispatch();
				m_Target_Sel.ReleaseDispatch();
				return TRUE;
			}
			m_Target_Sel.ReleaseDispatch();
		}
	}
	CATCH(CException, e)
	{
		e->Delete();
	}
	END_CATCH
	return FALSE;


}
BOOL CyzWordOperator::SetSelectTxtColor_New(DWORD clr,int Pos,BOOL bBOld,int chars)
{
	if(m_pWordApp==NULL)
		return FALSE;
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	TRY
	{
		if(m_Target_Sel.m_lpDispatch!=NULL)
		{
		/*	_Document myDoc; 
			myDoc=m_pWordApp->GetActiveDocument();
			COleVariant CStart((long)Pos),CEnd((long)(Pos+chars));
			Range rang=myDoc.Range(CStart,CEnd);
			//m_Target_Sel=m_pWordApp->GetSelection();
			_Font m_wdFt(rang.GetFont()); 
			m_wdFt.SetColor(clr);
			m_wdFt.SetBold(bBOld);
			rang.SetFont(m_wdFt.DetachDispatch());
			//rang.ReleaseDispatch();
			m_Target_Sel.SetRange(Pos,Pos);
			//myDoc.ReleaseDispatch();
			//m_Target_Sel.ReleaseDispatch();
			//m_Target_Sel.SetStart(Pos);
		*/
			m_Target_Sel.SetStart(Pos);
			m_Target_Sel.SetEnd(Pos+chars);				
			_Font m_wdFt(m_Target_Sel.GetFont()); 
			m_wdFt.SetColor(clr);
			m_Target_Sel.SetFont(m_wdFt.DetachDispatch());
			m_Target_Sel.SetStart(Pos);
			m_Target_Sel.SetEnd(Pos);
			//m_Target_Sel.ReleaseDispatch();
			return TRUE;
		}
	}
	CATCH(CException, e)
	{
		e->Delete();
	}
	END_CATCH
	return FALSE;


}
void CyzWordOperator::FooteSwitch()
{
	if(m_pWordApp==NULL)
		return ;
	long wdPrintView=3,wdWebView=6,wdPrintPreview=4,wdSeekFootnotes=7,wdPaneFootnotes=7;
	TRY
	{
		_Document ActiveDoc=m_pWordApp->GetActiveDocument();
		Footnotes footns=ActiveDoc.GetFootnotes();
		//是否存在脚注
		if(footns.GetCount()<=0)
			return;
		Window ActiveWin=m_pWordApp->GetActiveWindow();//激活的窗口
		View ActiveView=ActiveWin.GetView();//激活的视图
		Pane pane=ActiveWin.GetActivePane();//视图中激活的窗格
		View view=pane.GetView();
		long type=ActiveWin.GetType();
		if(type==wdPrintView||type==wdWebView||type==wdPrintPreview)
			ActiveView.SetSeekView(wdSeekFootnotes);
		else
			ActiveView.SetSplitSpecial(wdPaneFootnotes);
		return ;
	}
	CATCH(CException, e)
	{
		e->Delete();
		return ;
	}
	END_CATCH
		return ;
}

BOOL CyzWordOperator::ResorteTxt(int iStart, int iEnd)
{
	if(m_pWordApp==NULL)
		return FALSE;

	TRY
	{
		Selection m_Target_Sel=m_pWordApp->GetSelection();
		
		//int iCStart,iCEnd;
		//iCStart=m_Target_Sel.GetStart();
		//iCEnd=m_Target_Sel.GetEnd();

		if(m_Target_Sel.m_lpDispatch!=NULL)
		{
			_Document myDoc; 
			myDoc=m_pWordApp->GetActiveDocument();
			COleVariant CStart((long)iStart),CEnd((long)iEnd);
			Range rang=myDoc.Range(CStart,CEnd);
			//m_Target_Sel=m_pWordApp->GetSelection();
			_Font m_wdFt(rang.GetFont()); 
			m_wdFt.SetColor(RGB(0,0,0));
			m_wdFt.SetBold(0);
			rang.SetFont(m_wdFt.DetachDispatch());
			//m_Target_Sel.SetStart(iStart);
			//m_Target_Sel.SetEnd(iStart);
			//myDoc.ReleaseDispatch();
			
			/*m_Target_Sel.SetStart(iStart);
			m_Target_Sel.SetEnd(iEnd);				
			_Font m_wdFt(m_Target_Sel.GetFont()); 
			m_wdFt.SetColor(RGB(0,0,0));
			m_wdFt.SetBold(0);
			m_Target_Sel.SetFont(m_wdFt.DetachDispatch());
			m_Target_Sel.SetStart(iStart);
			m_Target_Sel.SetEnd(iStart);
			*/
			return TRUE;
		}


	}
	CATCH(CException, e)
	{
		e->Delete();
		return FALSE;
	}
	END_CATCH
		return FALSE;
}
/*在指定范围内查找给定字符keyword存在的位置，
找到，则返回开始位置iStart，未找到返回-1
*/

void CyzWordOperator::SearchKeyword(CString &keyword, int &iStart, int iEnd)
{
	if(m_pWordApp==NULL)
		return;

		
	Selection m_Target_Sel=m_pWordApp->GetSelection();

	m_Target_Sel.SetStart(iStart);
	//从头开始
	Find w_Find=m_Target_Sel.GetFind();
	
	//m_Target_Sel.ClearFormatting();//清除查找对象
	COleVariant FindText(keyword);
	COleVariant MatchCase((short)0), MatchWholeWord((short)0), MatchWildcards((short)0);
	COleVariant MatchSoundsLike((short)0), MatchAllWordForms((short)0),Forward((short)1),Wrap((short)1);
	COleVariant Format((short)0),ReplaceWith(""), Replace((short)0), MatchKashida((short)0),MatchDiacritics((short)0);
	COleVariant MatchAlefHamza((short)0), MatchControl((short)0);
	int ns=0;
	if(w_Find.Execute(FindText,MatchCase, MatchWholeWord, MatchWildcards,MatchSoundsLike, 
		MatchAllWordForms,Forward,Wrap,Format,ReplaceWith, Replace, MatchKashida,MatchDiacritics,MatchAlefHamza,
		MatchControl))
	{
		long lstart=m_Target_Sel.GetStart();
		if(iEnd>1)
		{
			long lsend=m_Target_Sel.GetEnd();
			if(lstart<iStart||lsend>iEnd)
			{
				m_Target_Sel.SetStart(iStart);
				m_Target_Sel.SetEnd(iStart);
				iStart=-1;
				return;
			}
			else if(lstart<lsend)
			{
				iStart=lstart;
				return;
			}
		}
		//m_Target_Sel.SetStart(lstart);
		//m_Target_Sel.SetEnd(lstart);
	}
	iStart=-1;
}
//取指开始定和结束点的文本
CString CyzWordOperator::GetStartAndEndText(int iStart, int iEnd)
{
	if(m_pWordApp==NULL)
		return "";
	
	Selection Sel=m_pWordApp->GetSelection();
	Sel.SetStart(iStart);
	Sel.SetEnd(iEnd);
	CString txt=Sel.GetText();
	return txt;

}

void CyzWordOperator::SetParagraphFormat()
{
	if(m_pWordApp==NULL)
		return;
	//FindWindowEx
	CWnd* pWnd=AfxGetMainWnd();
	//Caption MsoDockTop  class MsoCommandBarDock
	//HWND hWnd=FindWindowEx(pWnd->m_hWnd,NULL,_T("MsoCommandBar"),NULL);//"_WwG"_T("Microsoft Word 文档")
	HWND hWnd=::FindWindow(_T("MsoCommandBar"),NULL);//"_WwG"_T("Microsoft Word 文档")
	//HWND hWnd=::FindWindow(_T("MsoCommandBarDock"),"MsoDockTop");//"_WwG"_T("Microsoft Word 文档")
	if(hWnd!=NULL)
	{
		//::ShowWindow(hWnd,SW_HIDE);
		SendMessage(hWnd,WM_CLOSE,0,0);//MsoCommandBar
	}


	HRESULT hr;
	OLECHAR FAR* szMethod[2];
	DISPID dispid[2];
	VARIANT vArgs[3];
	DISPPARAMS dp;
/*	
	COleDispatchDriver CommandBar=m_pWordApp->GetCommandBars();
	szMethod[0]=OLESTR("Item");
	szMethod[1]=OLESTR("Visible");
	//查询Web条接口
	hr = CommandBar.m_lpDispatch->GetIDsOfNames(IID_NULL, szMethod, 1, 
		LOCALE_USER_DEFAULT, dispid);
	if(hr!=S_OK)
	{
		AfxMessageBox("error");
	}
	dp.cArgs = 2;
	
	dp.cNamedArgs = 2;
	dp.rgvarg = vArgs;
	dp.rgdispidNamedArgs=&(dispid[0]);  
	
	vArgs[2].vt = VT_BOOL;
	vArgs[2].iVal = 0;     //DoNotSetAsSysDefault = 1
	vArgs[1].vt = VT_BSTR;
	vArgs[1].bstrVal = ::SysAllocString(OLESTR("Visible"));

	vArgs[0].vt = VT_BSTR;
	vArgs[0].bstrVal = ::SysAllocString(OLESTR("Web"));
	//NOTE: You should replace "DeleteAllCommentsInDoc" in the line 
	//above with the name of a printer installed on your system.
	
	hr = CommandBar.m_lpDispatch->Invoke(dispid[0], IID_NULL, 
		LOCALE_USER_DEFAULT,DISPATCH_METHOD, &dp, NULL, NULL, NULL);
	
	::SysFreeString(vArgs[0].bstrVal);
	::SysFreeString(vArgs[1].bstrVal);
*/
	/*With ActiveWindow.View
	.ShowRevisionsAndComments = False
	.RevisionsView = wdRevisionsViewOriginal
    End With
	*/
	//删除所有的批注
	COleDispatchDriver WordBasic=m_pWordApp->GetWordBasic();
	
	//Retrieve the DISPIDs for the function as well as two of its named
	//arguments, Printer and DoNotSetAsSysDefault
	//查询接口
	szMethod[0]=OLESTR("DeleteAllCommentsInDoc"); //method name
	hr = WordBasic.m_lpDispatch->GetIDsOfNames(IID_NULL, szMethod, 1, 
		LOCALE_USER_DEFAULT, dispid);
	//Invoke the DeleteAllCommentsInDoc function using named arguments.
	//VARIANT vArgs[2];
	//DISPPARAMS dp;
	dp.cArgs = 0;
	
	dp.cNamedArgs = 0;
	dp.rgvarg = vArgs;
	dp.rgdispidNamedArgs=&(dispid[0]);  
	
	vArgs[1].vt = VT_I2;
	vArgs[1].iVal = 1;     //DoNotSetAsSysDefault = 1
	vArgs[0].vt = VT_BSTR;
	vArgs[0].bstrVal = ::SysAllocString(OLESTR("DeleteAllCommentsInDoc"));
	//NOTE: You should replace "DeleteAllCommentsInDoc" in the line 
	//above with the name of a printer installed on your system.
	
	hr = WordBasic.m_lpDispatch->Invoke(dispid[0], IID_NULL, 
		LOCALE_USER_DEFAULT,DISPATCH_METHOD, &dp, NULL, NULL, NULL);
	
	::SysFreeString(vArgs[0].bstrVal);
	//下面注释掉为不显示批注
	//Window pWindow=m_pWordApp->GetActiveWindow();
	
	//View pView=pWindow.GetView();
	//pView.SetShowRevisionsAndComments(0);
	//pView.SetRevisionsView(1);
	
	COleVariant Alignment((short)0), Leader((short)0);

	//_Document Doc=m_pWordApp->GetActiveDocument();

	
	Selection Sel=m_pWordApp->GetSelection();
	Sel.WholeStory();
	PageSetup pageSetup=Sel.GetPageSetup();
	pageSetup.SetLeftMargin(10.0);
    pageSetup.SetRightMargin(3.0);
	pageSetup.SetTopMargin(1.0);
	pageSetup.SetBottomMargin(1.0);
	//去掉项目符
	//Selection.Range.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph
	Range rngs=Sel.GetRange();
//	ListFormat lstfmt=rngs.GetListFormat();
//	COleVariant NumberType((short)1);//wdNumberParagraph
//	lstfmt.RemoveNumbers(NumberType);
	
	Sel.ClearFormatting();
	
	
	_Font font=Sel.GetFont();
	font.SetName("宋体");
	font.SetNameAscii("Times New Roman");
	font.SetNameOther("Times New Roman");
	font.SetSize(10.5);//5号
	font.SetSpacing(0);

	_ParagraphFormat Pgf=Sel.GetParagraphFormat();
	//Sel.ClearFormatting();
	Pgf.SetLeftIndent(0);
	Pgf.SetRightIndent(0);
	Pgf.SetSpaceBefore(0);
	Pgf.SetSpaceBeforeAuto(0);
	Pgf.SetSpaceAfter(0);
	Pgf.SetSpaceAfterAuto(0);
//	Pgf.SetLineSpacingRule(0);//1);//1.5倍行距,0单位行距
	Pgf.SetAlignment(3);//对齐方式
	Pgf.SetWidowControl(0);
	Pgf.SetKeepWithNext(0);
	Pgf.SetKeepTogether(0);

	
	Pgf.SetPageBreakBefore(0);
	Pgf.SetNoLineNumber(0);
	long True=-1;
	Pgf.SetHyphenation(True);//支持英文换行.SetHyphenation(1);用字连接符
	Pgf.SetFirstLineIndent(0);
	Pgf.SetOutlineLevel(10);
	Pgf.SetCharacterUnitLeftIndent(0);
	
	//.CharacterUnitRightIndent = 0
	Pgf.SetCharacterUnitRightIndent(0);
	//        .CharacterUnitFirstLineIndent = 0
	Pgf.SetCharacterUnitFirstLineIndent(0);
	//        .LineUnitBefore = 0
	Pgf.SetLineUnitBefore(0);
	Pgf.SetWordWrap(True);
	
	//        .LineUnitAfter = 0
	Pgf.SetLineUnitAfter(0);
	//        .AutoAdjustRightIndent = True
	Pgf.SetAutoAdjustRightIndent(True);
	//        .DisableLineHeightGrid = False
	Pgf.SetDisableLineHeightGrid(0);
	//        .FarEastLineBreakControl = True
	Pgf.SetFarEastLineBreakControl(True);
	
	//        .WordWrap = True
	//	Pgf.SetWordWrap(1);
	//        .HangingPunctuation = True
	Pgf.SetHangingPunctuation(True);
	//       .HalfWidthPunctuationOnTopOfLine = False
	Pgf.SetHalfWidthPunctuationOnTopOfLine(0);
	//        .AddSpaceBetweenFarEastAndAlpha = True
	Pgf.SetAddSpaceBetweenFarEastAndAlpha(True);
	//        .AddSpaceBetweenFarEastAndDigit = True
	Pgf.SetAddSpaceBetweenFarEastAndDigit(True);
	//        .BaseLineAlignment = wdBaselineAlignAuto
	Pgf.SetLineSpacingRule(1);//标准字距
	Pgf.SetBaseLineAlignment(4);
}

void CyzWordOperator::SetTabPos(float pos)
{
	if(m_pWordApp==NULL)
		return;
	//Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(7.3), _
    //Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
	COleVariant Alignment((short)0), Leader((short)0);

	Selection Sel=m_pWordApp->GetSelection();
	Sel.WholeStory();
	_ParagraphFormat Pgf=Sel.GetParagraphFormat();
	TabStops tabs=Pgf.GetTabStops();//SetTabStops().Add();
	tabs.ClearAll();
	tabs.Add(pos,Alignment,Leader);
}

BOOL CyzWordOperator::Replace_Space_Tab()
{
	if(m_pWordApp==NULL)
		return FALSE;

	Selection Sel=m_pWordApp->GetSelection();
	Sel.WholeStory();
	CString txt=Sel.GetText();
	BYTE Kh[3]="（";
	BYTE CVal,CBeVal=0,BeVal=0,AfVal=0,Val=0,CBeSpace=0;
	const BYTE Space=' ';
	int i=0;
	int StrLen=txt.GetLength();
	if(StrLen<=0)
		return FALSE;
	long lStart=-1,lEnd;
	int charp=1;
	while(i<StrLen)
	{
		CVal=txt.GetAt(i);
		if(CVal==Space)
		{
			if(lStart==-1 && BeVal!=Space)
			{
				lStart=charp,lEnd=charp;
				CBeSpace=Val;//保留空格前字符
			}
			else
				lEnd=charp;
			BeVal=Space;
			i++;
		}
		else
		{
			if(BeVal==Space)
			{
				if((lEnd-lStart)>2 && (CBeSpace!='(' && CBeSpace!=Kh[0]))
				{
					Sel.SetStart(lStart-1);
					Sel.SetEnd(lEnd);
					Sel.TypeText("\t");
					return TRUE;
				}
				lStart=-1;
				CBeVal=0;
				CBeSpace=0;

			}
			
			if(CVal>0xA0)
				i+=2;
			else 
				i++;
		}
		charp++;
		Val=CVal;
		BeVal=CVal;
	}
	return FALSE;
}

BOOL CyzWordOperator::FindReplace(CString Keyword, CString replace,short reps)
{
	if(m_pWordApp==NULL)
		return FALSE;

		
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	//从头开始
	Find w_Find=m_Target_Sel.GetFind();
	w_Find.ClearFormatting();
	//Replacement ReplaceObj=w_Find.GetReplacement();
	//ReplaceObj.ClearFormatting();
	//ReplaceObj.SetText(replace);
	COleVariant FindText(Keyword);
	COleVariant MatchCase((short)1), MatchWholeWord((short)0), MatchWildcards((short)0);
	COleVariant MatchSoundsLike((short)0), MatchAllWordForms((short)0),Forward((short)1),Wrap((short)1);//不回转
	COleVariant Format((short)0), MatchKashida((short)0),MatchDiacritics((short)0);
	COleVariant MatchAlefHamza((short)0), MatchControl((short)0);
	COleVariant Replace((short)reps);//2-全部替换，1-替换一处，0不替换
	COleVariant ReplaceWith(replace);
	int ns=0;
	if(w_Find.Execute(FindText,MatchCase, MatchWholeWord, MatchWildcards,MatchSoundsLike, 
		MatchAllWordForms,Forward,Wrap,Format,ReplaceWith, Replace, MatchKashida,MatchDiacritics,MatchAlefHamza,
		MatchControl))
	{
		//long lstart=m_Target_Sel.GetStart();
		//m_Target_Sel.SetStart(lstart);
		//m_Target_Sel.SetEnd(lstart);
		return TRUE;
	}
	return FALSE;
}

BOOL CyzWordOperator::FindReplace(CString Keyword, char Space)
{
	if(m_pWordApp==NULL)
		return FALSE;

		
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	//从头开始
	Find w_Find=m_Target_Sel.GetFind();
	w_Find.ClearFormatting();
	//Keyword="^p";//换段符
	COleVariant FindText(Keyword);
	COleVariant MatchCase((short)1), MatchWholeWord((short)0), MatchWildcards((short)0);
	COleVariant MatchSoundsLike((short)0), MatchAllWordForms((short)0),Forward((short)1),Wrap((short)0);////不回转
	COleVariant Format((short)0), MatchKashida((short)0),MatchDiacritics((short)0);
	COleVariant MatchAlefHamza((short)0), MatchControl((short)0);
	COleVariant Replace((short)0);//2-全部替换，1-替换一处，0不替换
	COleVariant ReplaceWith("");
	COleVariant Direction((short)1);//开方向

	int ns=0;
	
	bool bReplace=false;

	while(w_Find.Execute(FindText,MatchCase, MatchWholeWord, MatchWildcards,MatchSoundsLike, 
		MatchAllWordForms,Forward,Wrap,Format,ReplaceWith, Replace, MatchKashida,MatchDiacritics,MatchAlefHamza,
		MatchControl))
	{
		long lstart=m_Target_Sel.GetStart();
		long lend=m_Target_Sel.GetEnd();
		Selection st=m_pWordApp->GetSelection();
		CString txt1,txt2;
		//前字符
		st.SetStart(lstart-1);
		st.SetEnd(lstart-1);
		txt1=st.GetText();
		//取后字符
		st.SetStart(lend+1);
		st.SetEnd(lend+2);
		txt2=st.GetText();
		//AfxMessageBox(txt1+"\n"+txt2);

		//回到原来选区
		m_Target_Sel.SetStart(lstart);
		m_Target_Sel.SetEnd(lend);
		BYTE* pChar1=(BYTE*)txt1.GetBuffer(10);
		BYTE* pChar2=(BYTE*)txt2.GetBuffer(10);
/*
		CString str;
		str.Format("%d(%d,%d)",*pChar2,lstart,lend);
			AfxMessageBox(str);
*/
		bReplace=false;
		if(*pChar2=='\r'||*pChar2=='\n'||*pChar2=='\t'||*pChar2=='\v'|| *pChar2==' ')
		{
			bReplace=true;
		}
		else if(*pChar1=='\r' ||*pChar1=='\t' ||*pChar1=='\n' ||*pChar1=='\v')
			bReplace=true;
		
		if(bReplace)
		{
			//执行替换
			Replace=COleVariant((short)1);//替换一处
			m_Target_Sel.Collapse(Direction);
			w_Find.Execute(FindText,MatchCase, MatchWholeWord, MatchWildcards,MatchSoundsLike, 
			MatchAllWordForms,Forward,Wrap,Format,ReplaceWith, Replace, MatchKashida,MatchDiacritics,MatchAlefHamza,
			MatchControl);
		}
		//恢复查找
		Replace=COleVariant((short)0);

	}
	return FALSE;
}



void CyzWordOperator::SetDocSave(BOOL bSave)
{
	if(m_pWordApp==NULL)
		return ;
	_Document myDoc; 
	myDoc=m_pWordApp->GetActiveDocument();
	myDoc.SetSaved(bSave);
}
//删除页眉
void CyzWordOperator::DeletePageHeader()
{
	if(m_pWordApp==NULL)
		return ;
//	long wdPrintView=3,wdWebView=6,wdPrintPreview=4,wdSeekFootnotes=7,wdPaneFootnotes=7;
	TRY
	{
		_Document ActiveDoc=m_pWordApp->GetActiveDocument();
		Window ActiveWin=m_pWordApp->GetActiveWindow();//激活的窗口
		View ActiveView=ActiveWin.GetView();//激活的视图
		Pane pane=ActiveWin.GetActivePane();//视图中激活的窗格
		View view=pane.GetView();
		//激活页眉
		ActiveView.SetSeekView(9);
		Selection m_Target_Sel=m_pWordApp->GetSelection();
		m_Target_Sel.WholeStory();
		long start=m_Target_Sel.GetStart();
		long end=m_Target_Sel.GetEnd();
		COleVariant wdCharacter((short)1), Count((short)1);//(end-start));
		m_Target_Sel.Delete(wdCharacter,Count);
		//激活主页面
		ActiveView.SetSeekView(0);

		return ;

	}
	CATCH(CException, e)
	{
		return ;
	}
	END_CATCH
		return ;

}

void CyzWordOperator::FindKeywordParam(CString szKeyword,CDWordArray & StartAry)
{
	StartAry.RemoveAll();//清空
	if(m_pWordApp==NULL)
		return;

		
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	int iStart=m_Target_Sel.GetStart();
	int iEnd=m_Target_Sel.GetEnd();
	if(iStart>=iEnd)
		return;

	Find w_Find=m_Target_Sel.GetFind();
	w_Find.ClearFormatting();//清除查找对象
	//_Font font=w_Find.GetFont();
	//font.SetColor(RGB(255,0,0));
	COleVariant FindText(szKeyword);
	COleVariant MatchCase((short)0), MatchWholeWord((short)0), MatchWildcards((short)0);
	COleVariant MatchSoundsLike((short)0), MatchAllWordForms((short)0),Forward((short)1),Wrap((short)2);
	COleVariant Format((short)0),ReplaceWith(""), Replace((short)0), MatchKashida((short)0),MatchDiacritics((short)0);
	COleVariant MatchAlefHamza((short)0), MatchControl((short)0);

	int ns=0;
	while(w_Find.Execute(FindText,MatchCase, MatchWholeWord, MatchWildcards,MatchSoundsLike, 
		MatchAllWordForms,Forward,Wrap,Format,ReplaceWith, Replace, MatchKashida,MatchDiacritics,MatchAlefHamza,
		MatchControl))
	{
		DWORD nstart=m_Target_Sel.GetEnd();//.GetStart();
		if(nstart>(DWORD)iEnd)
			break;
		StartAry.Add(nstart);
	}
//	m_Target_Sel.SetStart(1);
//	m_Target_Sel.SetEnd(1);
}

BOOL CyzWordOperator::Select_Seg(int iStart, int iEnd)
{
	if(m_pWordApp==NULL)
		return FALSE;
//	if(m_pWordDoc)
//		m_pWordDoc->Activate();
	if(iStart>iEnd)
	{
		int ls=iStart;
		iStart=iEnd;
		iEnd=ls;
	}
	if(iEnd-iStart==0)
		return FALSE;
	TRY
	{
		Selection m_Target_Sel=m_pWordApp->GetSelection();
		m_Target_Sel.SetStart(iStart);
		m_Target_Sel.SetEnd(iEnd);
		m_Target_Sel.Select();
		int iStart_b=iStart;
		int iEnd_b=iEnd;
		CString str=m_Target_Sel.GetText();
		//if(str.GetLength()<=0)
			//return FALSE;
		int iLen=str.GetLength();
		char* pStr=str.LockBuffer();
		char mh[]="：　";
		int i=0;
		int nChars=0;
		//过滤掉开始位置的换行符,全角冒号和全角空格
		do
		{
			if(pStr[i]==32||pStr[i]=='\r'||pStr[i]=='\n'||pStr[i]==':')
			{
				i++;
				nChars++;
			}
			else if(((unsigned char)pStr[i])>0xA0)
			{
				if(pStr[i]==mh[0]&&pStr[i+1]==mh[1])
					i+=2,nChars++;
				else if(pStr[i]==mh[2]&&pStr[i+1]==mh[3])
					i+=2,nChars++;
				else
					break;
			}
			else
				break;

		}while(i<iLen);
		str.UnlockBuffer();

		//过滤掉后部无用字符
		int nEChars=0;
		int nEnter=0;
		int nEnter_f=0;
		//str.Find("\r\n\r\n");
		i=iLen-1;

		while(i>0)
		{
			if(pStr[i]==' '||pStr[i]=='\r'||pStr[i]=='\n'||pStr[i]==':')
			{
				if(pStr[i]=='\r')
				{
					nEnter++;
					nEnter_f=nEChars;
				}
				i--;
				nEChars++;
			}
			else if(((unsigned char)pStr[i])>0xA0)
			{
				if(pStr[i]==mh[1]&&pStr[i-1]==mh[0])
					i-=2,nEChars++;
				else if(pStr[i]==mh[3]&&pStr[i-1]==mh[2])
					i-=2,nEChars++;
				else
					break;
			}
			else
				break;

		}
		if(nChars>0)
			iStart+=nChars;
		if(nEChars>0 && nEnter>0)
		{
			iEnd-=nEnter_f;
		}
		if(iStart<iEnd)
		{
			m_Target_Sel.SetStart(iStart);
			m_Target_Sel.SetEnd(iEnd);
			m_Target_Sel.Select();
		}
		else
		{
			m_Target_Sel.SetStart(iStart_b);
			m_Target_Sel.SetEnd(iEnd_b);
			m_Target_Sel.Select();
		}

	}
	CATCH(CException, e)
	{
		e->Delete();
		return FALSE;
	}
	END_CATCH
		return TRUE;
}
//转换带下划线的空格
void CyzWordOperator::Replace_UnderLine(CString FindStr,CString ReplStr)
{
	if(m_pWordApp==NULL)
		return;// FALSE;
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	//从头开始
	Find w_Find=m_Target_Sel.GetFind();
	w_Find.ClearFormatting();
	_Font find_Font=w_Find.GetFont();
	find_Font.SetUnderline(1);//单下划线
	//替换字体格式
	Replacement w_Replace=w_Find.GetReplacement();
	w_Replace.ClearFormatting();
	_Font font=w_Replace.GetFont();
	font.SetUnderline(0);//无下划线
	
	//w_Find.SetFormat(TRUE);

	COleVariant FindText(FindStr);
	COleVariant MatchCase((short)1), MatchWholeWord((short)0), MatchWildcards((short)0);
	COleVariant MatchSoundsLike((short)0), MatchAllWordForms((short)0),Forward((short)1),Wrap((short)1);//不回转
	COleVariant Format((short)1), MatchKashida((short)0),MatchDiacritics((short)0);
	COleVariant MatchAlefHamza((short)0), MatchControl((short)0);
	COleVariant Replace((short)2);//2-全部替换，1-替换一处，0不替换
	COleVariant ReplaceWith(ReplStr);
	int ns=0;
	if(w_Find.Execute(FindText,MatchCase, MatchWholeWord, MatchWildcards,MatchSoundsLike, 
		MatchAllWordForms,Forward,Wrap,Format,ReplaceWith, Replace, MatchKashida,MatchDiacritics,MatchAlefHamza,
		MatchControl))
	{
		//long lstart=m_Target_Sel.GetStart();
		//m_Target_Sel.SetStart(lstart);
		//m_Target_Sel.SetEnd(lstart);
		return;// TRUE;
	}
	//return FALSE;
}
//设置图形居中
void CyzWordOperator::SetGraphCenter()
{
	if(m_pWordApp==NULL)
		return;// FALSE;
	Selection m_Target_Sel=m_pWordApp->GetSelection();


	int iStart,iEnd;
	iStart=m_Target_Sel.GetStart();
	iEnd=m_Target_Sel.GetEnd();
	m_Target_Sel.SetStart(iStart+1);
	m_Target_Sel.SetEnd(iEnd-1);
	m_Target_Sel.Select();
	
	//Paragraphs Pgs=m_Target_Sel.GetParagraphs();
	//long iPgs=Pgs.GetCount();//取段数
	_ParagraphFormat Pgf=m_Target_Sel.GetParagraphFormat();
	Pgf.SetAlignment(1);//对齐方式1-居中
	
	m_Target_Sel.SetStart(iStart);
	m_Target_Sel.SetEnd(iEnd);
	m_Target_Sel.Select();

}

BOOL CyzWordOperator::UserFind(CString szKeyword,int& iStart)
{
	if(m_pWordApp==NULL)
		return FALSE;
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	if(iStart>=0)
		m_Target_Sel.SetStart(iStart);
	else
		return FALSE;
	Find w_Find=m_Target_Sel.GetFind();
	w_Find.ClearFormatting();//清除查找对象
	//m_Target_Sel.ClearFormatting();
	COleVariant FindText(szKeyword);
	COleVariant MatchCase((short)0), MatchWholeWord((short)0), MatchWildcards((short)0);
	COleVariant MatchSoundsLike((short)0), MatchAllWordForms((short)0),Forward((short)1),Wrap((short)1);
	COleVariant Format((short)0),ReplaceWith(""), Replace((short)0), MatchKashida((short)0),MatchDiacritics((short)0);
	COleVariant MatchAlefHamza((short)0), MatchControl((short)0);
	BOOL bFind= w_Find.Execute(FindText,MatchCase, MatchWholeWord, MatchWildcards,MatchSoundsLike, 
		MatchAllWordForms,Forward,Wrap,Format,ReplaceWith, Replace, MatchKashida,MatchDiacritics,MatchAlefHamza,
		MatchControl);
	m_Target_Sel=m_pWordApp->GetSelection();
	iStart=m_Target_Sel.GetEnd();
	return bFind;

}

void CyzWordOperator::GetWholeStory(int &iStart, int &iEnd)
{
	if(m_pWordApp==NULL)
		return ;
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	m_Target_Sel.WholeStory();
	iStart=m_Target_Sel.GetStart();
	iEnd=m_Target_Sel.GetEnd();
	m_Target_Sel.SetStart(0);
}

_Document CWordHebin::CurDoc=NULL;
_Document CWordHebin::CurDaAn_Doc=NULL;//答案
_Document CWordHebin::CurDaAn_Txt=NULL;//答案
_Document CWordHebin::CurSTDA_Doc=NULL;//试案

CWordHebin::CWordHebin()
{
	m_nCloseApp=TRUE;//默认关闭
	m_nDocs=-1;
	m_TargetWord_App=NULL;
}

CWordHebin::~CWordHebin()
{
	if(m_TargetWord_App)
	{
		//退出
		m_TargetWord_App->m_bAutoRelease=TRUE;
		VARIANT vt ;
		vt.vt =VT_ERROR;
		vt.scode =DISP_E_PARAMNOTFOUND;
		
		VARIANT v;
		v.vt =VT_BOOL;
		v.boolVal =VARIANT_FALSE;
		if (m_nCloseApp)
		{
			m_TargetWord_App->Quit(&v,&vt,&vt);
			m_TargetWord_App->DetachDispatch();
			m_TargetWord_App->ReleaseDispatch();
		}
		else
			m_TargetWord_App->ReleaseDispatch();
		delete m_TargetWord_App;
	}
}


BOOL CWordHebin::CreateWordApp(int nDocs)
{
	if(::OpenClipboard(NULL))
	{
		::EmptyClipboard();
		::CloseClipboard();
	}
	m_nDocs=nDocs;
	switch(nDocs)
	{
	case 1:
		m_NoDocment[1]=&CurDoc;
		break;
	case 2:
		m_NoDocment[1]=&CurDoc;
		m_NoDocment[2]=&CurDaAn_Doc;
		break;
	case 3:
		m_NoDocment[1]=&CurDoc;
		m_NoDocment[2]=&CurDaAn_Doc;
		m_NoDocment[3]=&CurSTDA_Doc;
		break;
	}
		
	
	m_TargetWord_App=new _Application;
	//创建Word2003应用对象
	//AfxMessageBox("Word.Application.11");
	if (!m_TargetWord_App->CreateDispatch("Word.Application", NULL))
	{
		AfxMessageBox("启动Word2003应用程序建失败!\r\n机器上可能未安装Office 2003系统程序!", MB_OK | MB_SETFOREGROUND); 
		return FALSE;
	}
	m_nCloseApp=TRUE;

	Options Op=m_TargetWord_App->GetOptions();
	Op.SetCheckSpellingAsYouType(false);//拼写
	Op.SetCheckGrammarAsYouType(false);//语法检查
	Op.SetCheckGrammarWithSpelling(false);//用拼写检查语法
	Op.SetIgnoreUppercase(true);//忽略大小写


	
	COleVariant vFalse((long)0),vTrue((long)1);	
	COleVariant Template("");//Normal");//E:\\ENNORMAL.DOT
	COleVariant NewTemplate((short)0),DocumentType((short)0),Visible((short)TRUE);
	
	//	m_TargetWord_App->SetVisible(TRUE);
	
	Documents m_Target_Docs=m_TargetWord_App->GetDocuments();
	TRY
	{
		CurDoc=m_Target_Docs.Add(&Template,&NewTemplate,&DocumentType, &Visible);
		if(nDocs>1)
		{
			CurDaAn_Doc=m_Target_Docs.Add(&Template,&NewTemplate,&DocumentType, &Visible);
			CurDaAn_Txt=m_Target_Docs.Add(&Template,&NewTemplate,&DocumentType, &Visible);
		}
		if(nDocs>2)
			CurSTDA_Doc=m_Target_Docs.Add(&Template,&NewTemplate,&DocumentType, &Visible);

	}
	CATCH(CException, e)
	{
		return FALSE;
	}
	END_CATCH
		return TRUE;
}
//保存文档
void CWordHebin::SaveDocument(CString DocName,int nDocNo)
{
	if(nDocNo==1)
	{
		if(CurDoc==NULL)
		{
			AfxMessageBox("文档不存在!");
			return;
		}
	}
	else if(nDocNo==2)
	{
		if(CurDaAn_Doc==NULL)
		{
			AfxMessageBox("文档不存在!");
			return;
		}
	}
	else if(nDocNo==3)
	{
		if(CurSTDA_Doc==NULL)
		{
			AfxMessageBox("文档不存在!");
			return;
		}
	}



	VARIANT vt ;
	vt.vt =VT_ERROR;
	vt.scode =DISP_E_PARAMNOTFOUND;
	
	VARIANT varFileName;
	VariantInit(&varFileName);
	varFileName.vt =VT_BSTR;


		if(!DocName.IsEmpty())
		{
			varFileName.bstrVal = _bstr_t(DocName);
			Documents m_Target_Docs=m_TargetWord_App->GetDocuments();
			//不保存并关闭文档
			COleVariant vtrue((short)true), vfalse((short)false);
			if(nDocNo==1)
			{
				CurDoc.SaveAs(&varFileName,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt);
				CurDoc.Close(&vfalse, &vtrue,&vfalse);
				CurDoc.ReleaseDispatch();
			}
			else if(nDocNo==2)
			{
				CurDaAn_Doc.SaveAs(&varFileName,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt);
				CurDaAn_Doc.Close(&vfalse, &vtrue,&vfalse);
				CurDaAn_Doc.ReleaseDispatch();
			}
			else if(nDocNo==3)
			{
				CurSTDA_Doc.SaveAs(&varFileName,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt,&vt);
				CurSTDA_Doc.Close(&vfalse, &vtrue,&vfalse);
				CurSTDA_Doc.ReleaseDispatch();
			}

			if(CurDaAn_Txt)
			{
				CurDaAn_Txt.Close(&vfalse, &vtrue,&vfalse);
				CurDaAn_Txt.ReleaseDispatch();
				CurDaAn_Txt=NULL;
			}

			//清除文件名对象
			VariantClear(&varFileName);
			

		}//end of dlg.domodal
		else
		{
			AfxMessageBox("没有存储成功!");
		}	
}


void CWordHebin::Insert_TypeText(CString txt, CString szFontName, int Alignment, float size,int nDocNo,DWORD Color,BOOL bBold)
{
	if(m_TargetWord_App==NULL)
		return ;
	if(nDocNo==1)
	{
		//CurDoc.Select();
		CurDoc.Activate();
	}
	else if(nDocNo==2)
	{
		//CurDaAn_Doc.Select();
		CurDaAn_Doc.Activate();
	}
	else if(nDocNo==3)
		CurSTDA_Doc.Activate();

	Selection Sel=m_TargetWord_App->GetSelection();
	_Font OldFont=Sel.GetFont();
	if(size>0.0)
	{
		if(bBold)
			OldFont.SetBold(1);
		else
			OldFont.SetBold(0);
		OldFont.SetSize(size);
		OldFont.SetName(szFontName);
		OldFont.SetNameAscii("Times New Roman");
		OldFont.SetColor(Color);
	}
	_ParagraphFormat ParagFmt=Sel.GetParagraphFormat();
	ParagFmt.SetAlignment(Alignment);//1-居中
	Sel.TypeText(txt);
	ParagFmt.SetAlignment(3);//3-左对齐

}

BOOL CWordHebin::Insert_DocFile(CString szDocName,int nDocNo)
{
	if(m_TargetWord_App==NULL)
		return FALSE;
	try
	{
		if(nDocNo==1)
		{
			//CurDoc.Select();
			CurDoc.Activate();
			Selection Sel=m_TargetWord_App->GetSelection();
			COleVariant vFalse((long)0),vTrue((long)1);
			COleVariant vNull("");
			Sel.InsertFile(szDocName,&vNull,vFalse,vFalse,vFalse);
		}
		else if(nDocNo==2)
		{
			//CurDaAn_Doc.Select();
			CurDaAn_Doc.Activate();
			Selection Sel=m_TargetWord_App->GetSelection();
			COleVariant vFalse((long)0),vTrue((long)1);
			COleVariant vNull("");
			Sel.InsertFile(szDocName,&vNull,vFalse,vFalse,vFalse);
		}
		else if(nDocNo==3)
		{
						//CurDaAn_Doc.Select();
			CurSTDA_Doc.Activate();
			Selection Sel=m_TargetWord_App->GetSelection();
			COleVariant vFalse((long)0),vTrue((long)1);
			COleVariant vNull("");
			Sel.InsertFile(szDocName,&vNull,vFalse,vFalse,vFalse);
		}

	}
	catch(...)
	{
		return FALSE;
	}
	return TRUE;

}
//将指定Word文档文件打开选中复制到剪粘板中
void CWordHebin::OnOpenDoc_Copy(CString szDocFileName)
{
	CurDaAn_Txt.Select();
	CurDaAn_Txt.Activate();

	Selection Sel=m_TargetWord_App->GetSelection();
	Sel.WholeStory();
	COleVariant vFalse((long)0),vTrue((long)1);
	COleVariant vNull("");
	Sel.InsertFile(szDocFileName,&vNull,vFalse,vFalse,vFalse);
	//将选中内容置入剪粘板
	Sel.WholeStory();
	Sel.Copy();
}

BOOL CWordHebin::InsertBreakPage()
{
	if(m_TargetWord_App==NULL)
		return FALSE;
	TRY
	{
		Selection m_Target_Sel=m_TargetWord_App->GetSelection();
		COleVariant wdPageBreak((long)7);
		m_Target_Sel.InsertBreak(&wdPageBreak);
	}
	CATCH(CException, e)
	{
		return FALSE;
	}
	END_CATCH
		return TRUE;

}

void CWordHebin::SetVisible(BOOL bVisible)
{
	if(m_TargetWord_App)
	{
		m_TargetWord_App->SetVisible(bVisible);
	}


}
//获取文档的总行数
int CyzWordOperator::GetDocument_CountLines(int ParamCode)
{
	if(m_pWordApp==NULL)
		return -1;
	int iVal=-1;
	//项目索引号
	enum 
	{
		Number_of_pages=14,//页数
		Number_of_words=15,//字符(单词数)
		Number_of_Characters=16,//字符数
		Number_of_Bytes=22,//字节数
		Number_of_Lines=23,//行数
		Number_of_Paragraphs=24,//段落数 
		wdPropertyCharsWSpaces=30//含空格的字符数
	};
	_Document ActDoc=m_pWordApp->GetActiveDocument();//取当前激活的文档
	//取文档属性接口
	LPDISPATCH lpDisp=ActDoc.GetBuiltInDocumentProperties();
	COleDispatchDriver Obj_Inface;
	COleDispatchDriver Lines_Inface;
	Obj_Inface.AttachDispatch(lpDisp);

    DISPID dispID;                    // Temporary dispid for use in OleDispatchDriver::InvokeHelper().
    DISPID dispID2;                   // Dispid for 'Value'.
    unsigned short *ucPtr;            // Temporary dispid for use in OleDispatchDriver::InvokeHelper().
    VARIANT vtResult;                 // Holds results from OleDispatchDriver:: InvokeHelper().
	VARIANT vtResult2;                // Holds result for 'Type'.
	BYTE *parmStr;                    // Holds parameter descriptions  for COleDispatchDriver:: InvokeHelper().
	VARIANT i;                        // integer;
//	VARIANT count;                    // integer;
	
	try
	{
		ucPtr = L"Item";  // Collection has an Item member.
		Obj_Inface.m_lpDispatch->GetIDsOfNames(IID_NULL,&ucPtr,1,LOCALE_USER_DEFAULT,&dispID);
		parmStr = (BYTE *)( VTS_VARIANT );
		
		i.vt = VT_I4;
		i.lVal=ParamCode;
		
		Obj_Inface.InvokeHelper(dispID,DISPATCH_METHOD | DISPATCH_PROPERTYGET,VT_VARIANT,(void *)&vtResult,
			parmStr,&COleVariant(i));

		Lines_Inface.AttachDispatch(vtResult.pdispVal);
		ucPtr = L"Value";  // Collection has a Value member.
		Lines_Inface.m_lpDispatch->GetIDsOfNames(IID_NULL,&ucPtr,1,LOCALE_USER_DEFAULT,&dispID2);
		Lines_Inface.InvokeHelper(dispID2,DISPATCH_METHOD |DISPATCH_PROPERTYGET,VT_VARIANT,(void *)&vtResult2,NULL);
		if(vtResult2.vt==VT_I4)
		{
			iVal=vtResult2.lVal;
		}
		Lines_Inface.ReleaseDispatch();
		Obj_Inface.ReleaseDispatch();		
	}
	catch(COleException *e)
	{
		e->Delete();
	}
	return iVal;
}
CString  CyzWordOperator::GetDocument_CountInfo()
{
	if(m_pWordApp==NULL)
		return CString ("");
	int iVal=-1;
	//项目索引号
	int Param_Code[]={14,15,16,22,23,24,30};
/*	enum {Number_of_pages=14,//页数
		Number_of_words=15,//字符(单词数)
		Number_of_Characters=16,//字符数
		Number_of_Bytes=22,//字节数
		Number_of_Lines=23,//行数
		Number_of_Paragraphs=24,//段落数 
		wdPropertyCharsWSpaces=30//含空格的字符数
	};
*/
	CString info,txt;
	for(int i=0;i<sizeof(Param_Code)/sizeof(int);i++)
	{
		int iVal=GetDocument_CountLines(Param_Code[i]);
		switch(Param_Code[i])
		{
		case 14:
			txt.Format("页数=%d,",iVal);
			break;
		case 15:
			txt.Format("字数和单词数=%d,",iVal);
			break;
		case 16:
			txt.Format("字符数=%d,",iVal);
			break;
		case 22:
			txt.Format("字节数=%d,",iVal);
			break;
		case 23:
			txt.Format("行数=%d,",iVal);
			break;
		case 24:
			txt.Format("段落数=%d,",iVal);
			break;
		case 30:
			txt.Format("含空格的字符数=%d,",iVal);
			break;
//		default:
//			txt.Format("段落数=%d,",iVal);
			break;
		}
		info+=txt;
	}
	return info;
}

//取文档的总行数
//以下为参考代码,摘自Microsoft开发支持
/*
int CyzWordOperator::GetDocument_CountLines()
{
	if(m_pWordApp==NULL)
		return -1;
	
	_Document ActDoc=m_pWordApp->GetActiveDocument();//取当前激活的文档
	
	
	LPDISPATCH lpDisp=ActDoc.GetBuiltInDocumentProperties();
	COleDispatchDriver rootDisp[64];  //监时对象数组.
	int curRootIndex = 0;             //索引 数组.
    DISPID dispID;                    // Temporary dispid for use in OleDispatchDriver::InvokeHelper().
    DISPID dispID2;                   // Dispid for 'Value'.
    unsigned short *ucPtr;            // Temporary dispid for use in OleDispatchDriver::InvokeHelper().
    VARIANT vtResult;                 // Holds results from OleDispatchDriver:: InvokeHelper().
	VARIANT vtResult2;                // Holds result for 'Type'.
	BYTE *parmStr;                    // Holds parameter descriptions  for COleDispatchDriver:: InvokeHelper().
	rootDisp[0].AttachDispatch(lpDisp);  // LPDISPATCH returned from GetBuiltInDocumentProperties.
	VARIANT i;                        // integer;
	VARIANT count;                    // integer;
	char buf[512];                    // General purpose message buffer.
	char buf2[512];
	
	ucPtr = L"Count";                 // Collections have a Count member.
	try
	{
		//从ActDoc.GetBuiltInDocumentProperties()返回的dispID中查找"Count"入口地址
        rootDisp[curRootIndex].m_lpDispatch->GetIDsOfNames(IID_NULL,&ucPtr,1,LOCALE_USER_DEFAULT,&dispID);
		//调用dispID对应的函数
        rootDisp[curRootIndex].InvokeHelper(dispID,DISPATCH_METHOD |DISPATCH_PROPERTYGET,VT_VARIANT,(void *)&vtResult,
			NULL);
		
        count = vtResult;  // Require a separate variable for loop limiter.
        // For i = 1 to count,
        // get the Item, Name & Value members of the collection.
        i.vt = VT_I4;
        for(i.lVal=1; i.lVal<=count.lVal; i.lVal++)
        {
			ucPtr = L"Item";  // Collection has an Item member.
			rootDisp[curRootIndex].m_lpDispatch->GetIDsOfNames(IID_NULL,&ucPtr,1,LOCALE_USER_DEFAULT,&dispID);
			
			parmStr = (BYTE *)( VTS_VARIANT );
			rootDisp[curRootIndex].InvokeHelper(dispID,DISPATCH_METHOD | DISPATCH_PROPERTYGET,VT_VARIANT,(void *)&vtResult,
				parmStr,&COleVariant(i));
			
			// Move to the next element of the array.
			// Get the Name member for the Item.
			rootDisp[++curRootIndex].AttachDispatch(vtResult.pdispVal);
			ucPtr = L"Name";  // Collection has a Name member
			rootDisp[curRootIndex].m_lpDispatch->GetIDsOfNames(IID_NULL,&ucPtr,1,LOCALE_USER_DEFAULT,&dispID);
			
			rootDisp[curRootIndex].InvokeHelper(dispID,DISPATCH_METHOD |DISPATCH_PROPERTYGET,VT_VARIANT,(void *)&vtResult,
				NULL);
			
			ucPtr = L"Value";  // Collection has a Value member.
			rootDisp[curRootIndex].m_lpDispatch->GetIDsOfNames(IID_NULL,&ucPtr,1,LOCALE_USER_DEFAULT,&dispID2);
			
			rootDisp[curRootIndex].InvokeHelper(dispID2,DISPATCH_METHOD |DISPATCH_PROPERTYGET,VT_VARIANT,(void *)&vtResult2,NULL);
Continue: // Come back here from Catch(COleDispatchException).
			
			rootDisp[curRootIndex--].ReleaseDispatch();
			
			// Initialize buf2 with representation of the value.
			
			switch(vtResult2.vt) // Type of property.
			{
			case VT_BSTR:
				sprintf(buf2, "%s", (CString)vtResult2.bstrVal);
				break;
			case VT_DATE:
				{
					COleDateTime codt(vtResult2.date);
					sprintf(buf2, "Time = %d:%02d, Date = %d/%d/%d",
						codt.GetHour(), codt.GetMinute(),
						codt.GetMonth(), codt.GetDay(), codt.GetYear()
						);
				}
				break;
			case VT_I4:
				sprintf(buf2, "%ld", vtResult2.lVal);
				break;
			default:
				sprintf(buf2, "not VT_BSTR, VT_DATE, or VT_I4");
			}  // End of Switch.
			
			sprintf(buf, "Item(%d).Name = %s, .Type = %d, .Value = %s\n",
                i.lVal, CString(vtResult.bstrVal), vtResult2.vt, buf2);

			AfxMessageBox(buf);
			
			// objRange.Collapse(COleVariant((long)0));  // Move insertion point
			// to end of the range.
			//objRange.InsertAfter(CString(buf));  // Insert after the insertion
			// point.
			
        } 
		}
		
		catch(COleException *e)
		{
			sprintf(buf, "COleException. SCODE: %08lx.", (long)e->m_sc);
			::MessageBox(NULL, buf, "COleException", MB_SETFOREGROUND | MB_OK);
		}
		
		catch(COleDispatchException *e)
		{
			if(vtResult2.vt ==VT_ERROR)
			{
				AfxMessageBox("Discarding vtResult2.VT_ERROR");
				
			}
			vtResult2.vt = VT_BSTR;
			vtResult2.bstrVal = L"Value not available";
			goto Continue;
		}
		
		catch(...)
		{
			MessageBox(NULL,"General Exception caught.", "Catch-All",MB_SETFOREGROUND | MB_OK);
		}
		
		
}
*/
//关闭Word文档的工具条
void CyzWordOperator::CloseWordToolBar(BOOL bAll)
{
	if(m_pWordApp==NULL)
		return ;
	//取当前激活的文档对象
	_Document Doc=m_pWordApp->GetActiveDocument();
	_CommandBars mybars;
	CommandBar CloseBar;
	mybars=Doc.GetCommandBars();
	int ns=mybars.GetCount();
	CString txt,ls;
//	txt.Format("工具条个数%d",ns);
//	AfxMessageBox(txt);
	//COleVariant CloseObj(barName);
	//CloseBar=mybars.GetItem(CloseObj);
	//CloseBar.SetVisible(FALSE);
	for(int i=1;i<=ns;i++)
	{
		try
		{
			CloseBar=mybars.GetItem(COleVariant((short)i));
			ls=CloseBar.GetName();
			if(ls.CompareNoCase("Ribbon")==0)
				CloseBar.SetHeight(0);
			//CloseBar.GetProperty()
			
			//AfxMessageBox(ls);
			if(bAll)
				CloseBar.SetVisible(FALSE);
			else if(ls.CompareNoCase("Standard")!=0)
				CloseBar.SetVisible(FALSE);
			else if(ls.CompareNoCase("Menu Bar")!=0)
				CloseBar.SetVisible(FALSE);
			else
				CloseBar.SetVisible(TRUE);
			CloseBar.ReleaseDispatch();
		}
		catch(CException* e)
		{
			e->Delete();
		}
	}
}

void CyzWordOperator::FunAreaMin()
{
	if(m_pWordApp==NULL)
		return ;
	//改变视图显示比例
	Window WordWin;
	try
	{
		WordWin.AttachDispatch(m_pWordApp->GetActiveWindow());
		WordWin.ToggleRibbon();
	}
	catch(CException* e)
	{
		e->Delete();
	}
}
CyzWordTable::CyzWordTable()
{

}

CyzWordTable::~CyzWordTable()
{

}
//向Word当前位置添加表格
Table CyzWordTable::AddTable(int nRows,int nCols)
{
	if(!m_pWordApp)
		return 0;
	_Document myDoc; 
	myDoc=m_pWordApp->GetActiveDocument();
	Selection Sel=m_pWordApp->GetSelection();
	Sel.TypeText("\r\n");//插入换符
	Range rngs=Sel.GetRange();

	Tables tabs=myDoc.GetTables();
	COleVariant DefaultTableBehavior((short)1), AutoFitBehavior((short)0);
	Table tab=tabs.Add(rngs,nRows,nCols,&DefaultTableBehavior,&AutoFitBehavior);
	tab.AutoFitBehavior(0);//整体自动调整表格
	//表格整体居中
	Rows rows=tab.GetRows();
	// wdAlignRowCenter=1
	rows.SetAlignment(1);//表格整体居中
	return tab;
}
//合并单元格
BOOL CyzWordTable::HebinCell(Table TableObj,CPoint start, CPoint end)
{
	if(!m_pWordApp)
		return 0;
	_Document myDoc; 
	myDoc=m_pWordApp->GetActiveDocument();
	try
	{
		Cell sCel=TableObj.Cell(start.y,start.x);
		sCel.Select();
		Selection Sel=m_pWordApp->GetSelection();
		long lStart=Sel.GetStart();
		Sel.GetRange();
		Cell eCel=TableObj.Cell(end.y,end.x);
		eCel.Select();
		Sel=m_pWordApp->GetSelection();
		long lend=Sel.GetEnd();//.GetStart();
		Sel.SetStart(lStart);
		Sel.SetEnd(lend);
		Cells cells=Sel.GetCells();
		cells.Merge();
		return TRUE;
	}
	catch(CException* e)
	{
		e->Delete();
		AfxMessageBox("指定的合并区域无效!");
	}
	return FALSE;
}
//设置单元格数据
void CyzWordTable::SetCellText(Table tab,CPoint x, CString txt)
{
	Cell sCel=tab.Cell(x.y,x.x);
	Range Sel=sCel.GetRange();
	Sel.SetText(txt);
}
//设置单元格对齐方式
void CyzWordTable::SetCellFormat(Table tab,CPoint cell,int nFormat)
{
	Cell sCel=tab.Cell(cell.y,cell.x);
	sCel.Select();
	Selection PSel=m_pWordApp->GetSelection();

	_ParagraphFormat Pgf=PSel.GetParagraphFormat();
	Pgf.SetLeftIndent(0);
	Pgf.SetRightIndent(0);
	Pgf.SetSpaceBefore(0);
	Pgf.SetSpaceBeforeAuto(0);
	Pgf.SetSpaceAfter(0);
	Pgf.SetSpaceAfterAuto(0);
	Pgf.SetLineSpacingRule(0);//单倍行距
	Pgf.SetAlignment(nFormat);//对齐方式
	Pgf.SetWidowControl(0);
	Pgf.SetKeepWithNext(0);
	Pgf.SetKeepTogether(0);
	
	Pgf.SetPageBreakBefore(0);
	Pgf.SetNoLineNumber(0);
	long True=-1;
	Pgf.SetHyphenation(True);//支持英文换行.SetHyphenation(1);用字连接符
	Pgf.SetFirstLineIndent(0);
	Pgf.SetOutlineLevel(10);
	Pgf.SetCharacterUnitLeftIndent(0);
	
	//        .CharacterUnitRightIndent = 0
	Pgf.SetCharacterUnitRightIndent(0);
	//        .CharacterUnitFirstLineIndent = 0
	Pgf.SetCharacterUnitFirstLineIndent(0);
	//        .LineUnitBefore = 0
	Pgf.SetLineUnitBefore(0);
	
	//        .LineUnitAfter = 0
	Pgf.SetLineUnitAfter(0);
	//        .AutoAdjustRightIndent = True
	Pgf.SetAutoAdjustRightIndent(True);
	//        .DisableLineHeightGrid = False
	Pgf.SetDisableLineHeightGrid(True);
	//        .FarEastLineBreakControl = True
	Pgf.SetFarEastLineBreakControl(True);
	
	//        .WordWrap = True
	//	Pgf.SetWordWrap(1);
	//        .HangingPunctuation = True
	Pgf.SetHangingPunctuation(True);
	//       .HalfWidthPunctuationOnTopOfLine = False
	Pgf.SetHalfWidthPunctuationOnTopOfLine(0);
	//        .AddSpaceBetweenFarEastAndAlpha = True
	Pgf.SetAddSpaceBetweenFarEastAndAlpha(True);
	//        .AddSpaceBetweenFarEastAndDigit = True
	Pgf.SetAddSpaceBetweenFarEastAndDigit(True);
	//        .BaseLineAlignment = wdBaselineAlignAuto
	Pgf.SetBaseLineAlignment(4);
	
}
//	Cell sCel=tab.Cell(x.y,x.x);
void CyzWordTable::SetBorders(Table TabObj, CPoint pt)
{
	Cell sCel=TabObj.Cell(pt.y,pt.x);
	sCel.Select();
	Selection Sel=m_pWordApp->GetSelection();
	Cells cells=Sel.GetCells();
	cells.GetBorders();
	Borders borders=Sel.GetBorders();
	Border border=borders.Item(2);//左边框
	border.SetLineStyle(0);//无
	border=borders.Item(1);//顶边框
	border.SetLineStyle(0);//无
	border=borders.Item(4);//右边框
	border.SetLineStyle(0);//无
}

void CyzWordTable::PasteCells(Table &tab, CPoint start, CPoint end)
{
	Cell sCel=tab.Cell(start.y,start.x);
	sCel.Select();
	Selection Sel=m_pWordApp->GetSelection();
	long ls=Sel.GetStart();
	sCel=tab.Cell(end.y,end.x);
	sCel.Select();

	Sel=m_pWordApp->GetSelection();
	long le=Sel.GetEnd();
	Sel.SetStart(ls);
	Sel.SetEnd(le);

	Sel.Paste();
	
	sCel=tab.Cell(end.y,end.x);
	sCel.Select();
	Sel=m_pWordApp->GetSelection();
	le=Sel.GetEnd();

	Sel.SetStart(ls);
	Sel.SetEnd(le);
	_Font font=Sel.GetFont();
	font.SetSize(9.0);

	Cells cells=Sel.GetCells();
	cells.SetVerticalAlignment(1);//纵向居中
	
	_ParagraphFormat Pgf=Sel.GetParagraphFormat();
	Pgf.SetDisableLineHeightGrid(-1);
	Pgf.SetAutoAdjustRightIndent(0);
	Sel.SetStart(le+1);
	Sel.SetEnd(le+1);

}

void CyzWordOperator::Set_ShowParagraphs(BOOL bShow)
{
	if(m_pWordApp==NULL)
		GetWordAppObj();
	if(m_pWordApp)
	{
		Window ActiveWindow=m_pWordApp->GetActiveWindow();
		View Viewa=ActiveWindow.GetView();
		Viewa.SetShowParagraphs(bShow);
		//Viewa.SetType(model);
	}

}

BOOL CyzWordOperator::IsColor(COLORREF color)
{
	if(m_pWordApp==NULL)
		return FALSE;

		
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	int iStart=m_Target_Sel.GetStart();
	int iEnd=m_Target_Sel.GetEnd();
//	if(iStart>=iEnd)
//		return FALSE;
	//m_Target_Sel.SetStart(iStart);
	//m_Target_Sel.SetEnd(iStart);
	Range rng=m_Target_Sel.GetRange();
	//m_Target_Sel=m_pWordApp->GetSelection();
	Find w_Find=rng.GetFind();
/*	
	_Font ch=rng.GetFont();
	long colr=ch.GetColor();
	if(ch.GetColor()==(long)color)
		return TRUE;
	return FALSE;
*/
	_Font font=w_Find.GetFont();
	w_Find.ClearFormatting();
	font.SetColor(color);

	COleVariant FindText("");
	COleVariant MatchCase((short)0), MatchWholeWord((short)0), MatchWildcards((short)0);
	COleVariant MatchSoundsLike((short)0), MatchAllWordForms((short)0),Forward((short)1),Wrap((short)0);
	COleVariant Format((short)1),ReplaceWith(""), Replace((short)0), MatchKashida((short)0),MatchDiacritics((short)0);
	COleVariant MatchAlefHamza((short)0), MatchControl((short)0);

	int ns=0;
	if(w_Find.Execute(FindText,MatchCase, MatchWholeWord, MatchWildcards,MatchSoundsLike, 
		MatchAllWordForms,Forward,Wrap,Format,ReplaceWith, Replace, MatchKashida,MatchDiacritics,MatchAlefHamza,
		MatchControl))
	{
		m_Target_Sel.SetStart(iStart);
		m_Target_Sel.SetEnd(iEnd);

		return TRUE;
	}
	return FALSE;
}
//取Word版本号
CString CyzWordOperator::GetWordVersion()
{
	CString szVersion;
	if(m_pWordApp)
		m_pWordApp->InvokeHelper(0x18, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&szVersion, NULL);
	return szVersion;
	_variant_t;
	//取Excel版本号
	//InvokeHelper(0x188, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&szVersion, NULL);
	//InvokeHelper(0x7de, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&szVersion, NULL);//取Powerpoint版本号

}

void CyzWordTable::SetColumnWide(Table &table, int Col, float wide)
{
	Selection Sel=m_pWordApp->GetSelection();

	Columns columns=table.GetColumns();
	int ns=columns.GetCount();
	Column column=columns.Item(Col);
	column.SetWidth(wide,0);
	//column.Select();

}

void CyzWordTable::SetColumnForat(Table tab, int Col, int nFormat)
{
	Columns columns=tab.GetColumns();
	int ns=columns.GetCount();
	Column column=columns.Item(Col);
	column.Select();
	Selection PSel=m_pWordApp->GetSelection();
	_ParagraphFormat Pgf=PSel.GetParagraphFormat();
	Pgf.SetAlignment(nFormat);//对齐方式
	//表格整体居中
	Rows rows=tab.GetRows();
	// wdAlignRowCenter=1
	rows.SetAlignment(1);//表格整体居中
	tab.AutoFitBehavior(1);//整体自动调整表格
	
}
//获取表格的结束位置
long CyzWordTable::GetEndPos(Table &table)
{
	Columns columns=table.GetColumns();
	int ns=columns.GetCount();
	Range range=table.GetRange();
	return range.GetEnd();
}

void CyzWordTable::SetRowFormat(Table &tab,int irow)
{
	Columns columns=tab.GetColumns();
/*	int ns=columns.GetCount();
	Column column=columns.Item(Col);
	column.Select();
	Selection PSel=m_pWordApp->GetSelection();
	_ParagraphFormat Pgf=PSel.GetParagraphFormat();
	Pgf.SetAlignment(nFormat);//对齐方式
*/
	//表格整体居中
	Rows rows=tab.GetRows();
	// wdAlignRowCenter=1
//	rows.SetAlignment(1);//表格整体居中
//	tab.AutoFitBehavior(1);//整体自动调整表格
	Row row=rows.Item(irow);
	row.Select();
	Selection PSel=m_pWordApp->GetSelection();
	_ParagraphFormat Pgf=PSel.GetParagraphFormat();
	Pgf.SetAlignment(1);//对齐方式
	Pgf.SetDisableLineHeightGrid(-1);//禁止网络对齐
	Pgf.SetAutoAdjustRightIndent(0);
	Pgf.SetLineSpacingRule(0);
	Cells cells=row.GetCells();
	cells.SetVerticalAlignment(1);//纵向居中
}

void CyzWordOperator::Search_StInfo_Shjuan(CString Keyword, CStringArray &InfoAry)
{
	InfoAry.RemoveAll();
	if(m_pWordApp==NULL)
		return ;
	m_StInfoInDoc.FreeMapTab();
	m_WordDoc.Activate();
	//取出总行数
	int nmax=GetWordDocLines();
	m_StInfoInDoc.m_Doc_Max_Line=nmax;
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	COleVariant What((short)3),Which((short)1),Continue((short)1),Name("");
	//插入光标移动到指定行
	COleVariant Unit((short)5),Extend((short)TRUE);
	COleVariant	Duan_Flag((short)4);//段标识
	
	if(m_pStParamTab)
		m_pStParamTab->RemoveAll();
	for(int is=1;is<=nmax;is++)
	{
		GotoLine(is);
		m_Target_Sel.EndOf(&Duan_Flag,&Extend);//移到段尾;
		DWORD pos=m_Target_Sel.GetStart();
		CString str=m_Target_Sel.GetText();
		//AfxMessageBox(str);
		//读取试题编号
		if(str.Find(Keyword,0)>=0)
		{
			CString str1,str2,str3,str4;
			int idx=str.Find(Keyword);
			int i;
			for(i=idx;i>=0&&i<str.GetLength();i++)
			{
				char chr=str.GetAt(i);
				if((chr>='0'&&chr<='9')||(chr>='a'&&chr<='z')||(chr>='A'&&chr<='Z'))
					str1+=chr;
				if(chr==']')
					break;
			}
			//取分值
			idx=str.Find(str1);//找出试题编号存在的位置
			str2=str.Mid(idx+str1.GetLength());
			str3.Empty();
			for(i=0;i<str2.GetLength();i++)
			{
				char chr=str2.GetAt(i);
				if((BYTE)chr>160)
					break;
				else if((chr>='0'&&chr<='9')||chr=='.')
					str3+=chr;
			}
			COleVariant Page=m_Target_Sel.GetInformation(1);//取当前页号//(10);
			//COleVariant Line=m_Target_Sel.GetInformation(10);//取当前行号//(10);
			str.Format("%d,%d,%s,%s",is,pos,str3,str1);//行号，页号，分值，试题编号
			InfoAry.Add(str);
			str.Format("%s:%s",str1,str3);//题号与分值
			//AfxMessageBox(str);
			m_StInfoInDoc.AddParamStartPos(pos,Page.intVal,is,str,m_pStParamTab);
		}
	}
	GotoLine(1);
}

CString CyzWordOperator::GetKeywordLineInfo(CString Keyword)
{
		if(m_pWordApp==NULL)
		return "";
	m_WordDoc.Activate();
	try
	{
		Selection m_Target_Sel=m_pWordApp->GetSelection();
		//从头开始
		Find w_Find=m_Target_Sel.GetFind();
		//m_Target_Sel.ClearFormatting();//清除查找对象
		COleVariant FindText(Keyword);
		COleVariant MatchCase((short)0), MatchWholeWord((short)0), MatchWildcards((short)0);
		COleVariant MatchSoundsLike((short)0), MatchAllWordForms((short)0),Forward((short)1),Wrap((short)1);
		COleVariant Format((short)0),ReplaceWith(""), Replace((short)0), MatchKashida((short)0),MatchDiacritics((short)0);
		COleVariant MatchAlefHamza((short)0), MatchControl((short)0);
		
		int ns=0;
		if(w_Find.Execute(FindText,MatchCase, MatchWholeWord, MatchWildcards,MatchSoundsLike, 
			MatchAllWordForms,Forward,Wrap,Format,ReplaceWith, Replace, MatchKashida,MatchDiacritics,MatchAlefHamza,
			MatchControl))
		{
			CString txt;
			COleVariant Uint(short(5));//wdLine;
			COleVariant Extend(short(0));//wdExtend
			m_Target_Sel.HomeKey(&Uint,&Extend);//移至行头
			Extend.intVal=1;
			m_Target_Sel.EndKey(&Uint,&Extend);//移到行尾
			txt=m_Target_Sel.GetText();
			return txt;
		}
	}
	catch(CException* e)
	{
		e->Delete();
	}
	return "";
}
//找当前位置后的字符
BOOL CyzWordOperator::Search_Char_Pos(CString find_Txt, int &Pos)
{
	if(m_pWordApp==NULL)
		return FALSE;

	try
	{
		Selection m_Target_Sel=m_pWordApp->GetSelection();
		//从头开始
		Find w_Find=m_Target_Sel.GetFind();
		COleVariant Text(find_Txt);
		COleVariant Replace((short)0);
		COleVariant Forward((short)1),Wrap((short)1), Format((short)0),MatchCase((short)0);

		COleVariant	MatchWholeWord((short)0), MatchWildcards((short)0);
		COleVariant MatchSoundsLike((short)0), MatchAllWordForms((short)0);
		COleVariant ReplaceWith(""),  MatchKashida((short)0),MatchDiacritics((short)0);
		COleVariant MatchAlefHamza((short)0), MatchControl((short)0);
		int ns=0;
		if(w_Find.Execute(Text,MatchCase, MatchWholeWord, MatchWildcards,MatchSoundsLike, 
			MatchAllWordForms,Forward,Wrap,Format,ReplaceWith, Replace, MatchKashida,MatchDiacritics,MatchAlefHamza,
			MatchControl))
		{
			Pos=m_Target_Sel.GetStart();
			m_Target_Sel.SetStart(Pos);
			m_Target_Sel.SetEnd(Pos);
			return TRUE;
		}
	}
	catch(CException* e)
	{
		e->Delete();
	}
	return FALSE;
}

void CWordHebin::MoveEnd()
{
	_Document myDoc; 
	myDoc=m_TargetWord_App->GetActiveDocument();
	Selection Sel=m_TargetWord_App->GetSelection();
	COleVariant Unit((short)6),Count((short)0);
	Sel.EndKey(Unit,Count);//.MoveEnd();
}

void CWordHebin::SetParagraphFormat()
{
	if(m_TargetWord_App==NULL)
		return;
	//FindWindowEx
	Selection Sel=m_TargetWord_App->GetSelection();
	_ParagraphFormat Pgf=Sel.GetParagraphFormat();
	//Sel.ClearFormatting();
	Pgf.SetLeftIndent(0);
	Pgf.SetRightIndent(0);
	Pgf.SetSpaceBefore(0);
	Pgf.SetSpaceBeforeAuto(0);
	Pgf.SetSpaceAfter(0);
	Pgf.SetSpaceAfterAuto(0);
	Pgf.SetLineSpacingRule(0);//1.5倍行距
	Pgf.SetAlignment(3);//对齐方式
	Pgf.SetWidowControl(0);
	Pgf.SetKeepWithNext(0);
	Pgf.SetKeepTogether(0);
	Pgf.SetPageBreakBefore(0);
	Pgf.SetNoLineNumber(0);
	long True=-1;
	Pgf.SetHyphenation(True);//支持英文换行.SetHyphenation(1);用字连接符
	Pgf.SetFirstLineIndent(0);
	Pgf.SetOutlineLevel(10);
	Pgf.SetCharacterUnitLeftIndent(0);
	
	//.CharacterUnitRightIndent = 0
	Pgf.SetCharacterUnitRightIndent(0);
	//        .CharacterUnitFirstLineIndent = 0
	Pgf.SetCharacterUnitFirstLineIndent(0);
	//        .LineUnitBefore = 0
	Pgf.SetLineUnitBefore(0);
	Pgf.SetWordWrap(True);
	
	//        .LineUnitAfter = 0
	Pgf.SetLineUnitAfter(0);
	//        .AutoAdjustRightIndent = True
	Pgf.SetAutoAdjustRightIndent(0);
	//        .DisableLineHeightGrid = False
	Pgf.SetDisableLineHeightGrid(True);
	//        .FarEastLineBreakControl = True
	Pgf.SetFarEastLineBreakControl(True);
	
	//        .WordWrap = True
	Pgf.SetWordWrap(True);//.SetWordWrap(1);
	//        .HangingPunctuation = True
	Pgf.SetHangingPunctuation(True);
	//       .HalfWidthPunctuationOnTopOfLine = False
	Pgf.SetHalfWidthPunctuationOnTopOfLine(0);
	//        .AddSpaceBetweenFarEastAndAlpha = True
	Pgf.SetAddSpaceBetweenFarEastAndAlpha(True);
	//        .AddSpaceBetweenFarEastAndDigit = True
	Pgf.SetAddSpaceBetweenFarEastAndDigit(True);
	//        .BaseLineAlignment = wdBaselineAlignAuto
	Pgf.SetBaseLineAlignment(4);
}

void CWordHebin::InsertTable(CyzWordTable &WordTable, int Rows, int Cols, CPtrArray &wides)
{
	if(!m_TargetWord_App)
		return ;
	_Document myDoc; 
	myDoc=m_TargetWord_App->GetActiveDocument();
	Selection Sel=m_TargetWord_App->GetSelection();

	Range rngs=Sel.GetRange();

	Tables tabs=myDoc.GetTables();
	COleVariant DefaultTableBehavior((short)1), AutoFitBehavior((short)0);
	WordTable.m_CurTable=tabs.Add(rngs,Rows,Cols,&DefaultTableBehavior,&AutoFitBehavior);
	Table& tab=WordTable.m_CurTable;
	Columns columns=tab.GetColumns();
	columns.GetCount();
	
	for(int i=0;i<wides.GetSize();i++)
	{
		Column column=columns.Item(1+i);
		column.SetWidth(*((float*)wides.GetAt(i)),0);
	}
}

void CyzWordOperator::SetMargin()
{
	if(!m_pWordApp)
		return ;
	//m_pWordApp->GetActiveWindow();
	Selection Sel=m_pWordApp->GetSelection();
	PageSetup pageSetup=Sel.GetPageSetup();
	pageSetup.SetLeftMargin(10.0);
    pageSetup.SetRightMargin(8.0);
	pageSetup.SetTopMargin(1.0);
	pageSetup.SetBottomMargin(1.0);

}

void CWordHebin::SetMargin()
{
	if(!m_TargetWord_App)
		return ;
	Selection Sel=m_TargetWord_App->GetSelection();
	PageSetup pageSetup=Sel.GetPageSetup();
	pageSetup.SetLeftMargin(10.0);
    pageSetup.SetRightMargin(8.0);
	pageSetup.SetTopMargin(1.0);
	pageSetup.SetBottomMargin(1.0);

}

void CyzWordOperator::SetDisplayRulers(BOOL bDisplay)
{
	if(!m_pWordApp)
		return ;
	Window WordWin;
	WordWin.AttachDispatch(m_pWordApp->GetActiveWindow());
	WordWin.SetDisplayRulers(bDisplay);

}

int CyzWordOperator::GetSelectRangeKeyPos(CString keyword, int iStart, int iEnd)
{
	if(m_pWordApp==NULL)
		return 0;
	
	Selection Sel=m_pWordApp->GetSelection();
	Sel.SetStart(iStart);
	Sel.SetEnd(iEnd);
	Sel.GetFind();
	Find w_Find=Sel.GetFind();
	w_Find.ClearFormatting();
	//m_Target_Sel.ClearFormatting();//清除查找对象
	COleVariant FindText(keyword);
	COleVariant MatchCase((short)1), MatchWholeWord((short)0), MatchWildcards((short)0);
	COleVariant MatchSoundsLike((short)0), MatchAllWordForms((short)0),Forward((short)0),Wrap((short)0);
	COleVariant Format((short)0),ReplaceWith(""), Replace((short)0), MatchKashida((short)0),MatchDiacritics((short)0);
	COleVariant MatchAlefHamza((short)0), MatchControl((short)0);

	int ns=0;
	if(w_Find.Execute(FindText,MatchCase, MatchWholeWord, MatchWildcards,MatchSoundsLike, 
		MatchAllWordForms,Forward,Wrap,Format,ReplaceWith, Replace, MatchKashida,MatchDiacritics,MatchAlefHamza,
		MatchControl))
	{
		DWORD nstart=Sel.GetStart();
		return nstart;
	}
	return 0;
}

void CyzWordOperator::SetHyperLink(CString strLink)
{
	//ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="000A", _SubAddress:=""
	Selection Sel=m_pWordApp->GetSelection();
	_Document myDoc; 
	myDoc=m_pWordApp->GetActiveDocument();
	Hyperlinks hyperlinks=myDoc.GetHyperlinks();
	COleVariant Address(strLink);
	COleVariant SubAddress("");
	COleVariant ScreenTip("");
	COleVariant TextToDisplay("");
	COleVariant Target("");
	//Add(LPDISPATCH Anchor, VARIANT* Address, VARIANT* SubAddress, VARIANT* ScreenTip, VARIANT* TextToDisplay, VARIANT* Target)
	hyperlinks.Add(Sel.GetRange(),Address,SubAddress,ScreenTip,TextToDisplay,Target);

}

CString CyzWordOperator::GetHyperLinksText()
{
	Selection Sel=m_pWordApp->GetSelection();
	_Document myDoc; 
	myDoc=m_pWordApp->GetActiveDocument();
	Hyperlinks hyperlinks=myDoc.GetHyperlinks();
	int links=hyperlinks.GetCount();
	if(links<1)
		return "";
	CString Address("");
	for(short i=1;i<=links;i++)
	{
		COleVariant item(i);
		Hyperlink  hylink=hyperlinks.Item(item);
		CString address=hylink.GetAddress();
		if(Address.IsEmpty())
			Address=address;
		else
			Address+="\n"+address;
	}
	return Address;

	//Add(LPDISPATCH Anchor, VARIANT* Address, VARIANT* SubAddress, VARIANT* ScreenTip, VARIANT* TextToDisplay, VARIANT* Target)

}

void CyzWordOperator::SetDisplayScreenTips(BOOL bShow)
{
	if(m_pWordApp)
	{
		Window window=m_pWordApp->GetActiveWindow();
		window.SetDisplayScreenTips(FALSE);
	}

}

void CyzWordOperator::CopyItem(CWnd* pWnd,UINT uPaseCmd,BOOL& bEquation)
{
	if(m_pWordApp==NULL)
		return ;
/*
		
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	m_Target_Sel.WholeStory();
	InlineShapes shapes=m_Target_Sel.GetInlineShapes();
	long count=shapes.GetCount();
	for(int i=0;i<count;i++)
	{
		InlineShape shapeObj=shapes.Item(i+1);
		shapeObj.Select();
		Range rang=m_Target_Sel.GetRange();
		rang.Copy();
		if(pWnd)
			pWnd->SendMessage(WM_COMMAND,uPaseCmd,0);
		m_Target_Sel.WholeStory();

	}
*/	
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	m_WordDoc.Activate();
	InlineShapes shapes=m_WordDoc.GetInlineShapes();
	long count=shapes.GetCount();
	for(int i=0;i<count;i++)
	{
		InlineShape shapeObj=shapes.Item(i+1);
		OLEFormat OleFormat=shapeObj.GetOLEFormat();
		CString classType=OleFormat.GetClassType();
		classType.MakeLower();
		if(classType.GetLength()>0 && classType.Find("equation")>=0)
			bEquation=TRUE;
		else
			bEquation=FALSE;
		
			shapeObj.Select();
			Range rang=m_Target_Sel.GetRange();
			rang.Copy();
			if(pWnd)
				pWnd->SendMessage(WM_COMMAND,uPaseCmd,0);
		

	}
}

void CyzWordOperator::ShowEditFlag(BOOL bShow)
{
	if(m_pWordApp)
	{
		Window window=m_pWordApp->GetActiveWindow();
		Pane pane=window.GetActivePane();
		View view=pane.GetView();//window.GetView();
		view.SetShowAll(bShow);
	}

}
//将数学公式替换为关键词，每次替换一个
BOOL CyzWordOperator::ReplaceEquation(CString Keyword)
{
	if(m_pWordApp==NULL)
		return FALSE;

	
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	m_WordDoc.Activate();
	InlineShapes shapes=m_WordDoc.GetInlineShapes();
	long count=shapes.GetCount();
	for(int i=0;i<count;i++)
	{
		InlineShape shapeObj=shapes.Item(i+1);
		OLEFormat OleFormat=shapeObj.GetOLEFormat();
		CString classType=OleFormat.GetClassType();
		classType.MakeLower();
		//if(classType.GetLength()>0 && classType.Find("equation")>=0)
		{
			//AfxMessageBox(classType);
			shapeObj.Select();
			Range rang=m_Target_Sel.GetRange();
			rang.SetText(Keyword);
			return TRUE;
		}
	}
	return FALSE;
}

void CyzWordOperator::ClearAll_HyperLink()
{
	if(m_pWordApp==NULL)
		return;

	m_WordDoc.Activate();
	Hyperlinks hyperlinks=m_WordDoc.GetHyperlinks();
	int links=hyperlinks.GetCount();
	if(links<1)
		return ;
	for(short i=links;i>0;i--)
	{
		COleVariant item(i);
		Hyperlink  hylink=hyperlinks.Item(item);
		hylink.Delete();
	}

}

int CyzWordOperator::GetTableInfo()
{
	if(m_pWordApp==NULL)
		return -1;
	//Selection m_Target_Sel=m_pWordApp->GetSelection();
	m_WordDoc.Activate();
	Tables tabs=m_WordDoc.GetTables();
	long lbs=tabs.GetCount();
	if(lbs<1)
		return 0;
	return lbs;
}
//读取单格信息
CString CyzWordOperator::GetTableCellText(int iTabNo, int row, int col,long& endpos)
{
	if(m_pWordApp==NULL)
		return "";
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	m_WordDoc.Activate();
	Tables tabs=m_WordDoc.GetTables();
	long lbs=tabs.GetCount();
	if(lbs<1)
		return "";
	//选定表
	Table table=tabs.Item(iTabNo);
	CString info;
	//选定行列
	Cell cell=table.Cell(row,col);
	cell.Select();
	Range rgn=cell.GetRange();				
	info=rgn.GetText();
	info.Replace("\r","");
	info.Replace("\n","");
	info.Replace("\007","");//去掉响铃符
	info.Replace("\001","[Object]");
	endpos=rgn.GetEnd();
	return info;
}

CString CyzWordOperator::GetCurParagraphInfo(CString StartTxt,CString EndTxt)
{
	if(m_pWordApp==NULL)
		return "";
	m_WordDoc.Activate();
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	//取当前插入位置
	Range Sel=m_Target_Sel.GetRange();
	Paragraphs duanluos=Sel.GetParagraphs();
	long count=duanluos.GetCount();
	
	Paragraph dl=duanluos.Item(1);
	COleVariant Count((long)1);
	CString Text(""),ls;
	do
	{
		Range dlrgn=dl.GetRange();
		ls=dlrgn.GetText();
		Words words=dlrgn.GetWords();
		long lws=words.GetCount();
		if(ls.Find(StartTxt)>=0)
			Text+=ls;
		else if(EndTxt.GetLength()>0 && ls.Find(EndTxt)>=0)
			break;
		else if(Text.GetLength()>0)
		{
			CString test=ls;
			test.Replace("\r","");
			test.Replace("\n","");
			test.Replace(" ","");
			if(test.GetLength()>0)
				Text+=ls;
		}
		dl=dl.Next(Count);//取下一段落
		if(dl.m_lpDispatch==NULL)
			break;
	}while(1);
	return Text;
}
//获Word绘图对象数
long CyzWordOperator::GetShapesGroupCount()
{
	m_WordDoc.Activate();
	Shapes sps=m_WordDoc.GetShapes();
	int ic=sps.GetCount();
/*	COleVariant no((short)1);
	Shape spa=sps.Item(no);
	CString name=spa.GetName();
*/
	return ic;
}

void CyzWordOperator::SetFontInfo(CString ZwName, CString XwName, float size, COLORREF color)
{
	if(m_pWordApp==NULL)
		return ;

	Selection m_Target_Sel=m_pWordApp->GetSelection();//.GetSelection();
	_Font m_wdFt = m_Target_Sel.GetFont();

	m_wdFt.SetSize(size);
	m_wdFt.SetName(ZwName);
	m_wdFt.SetNameAscii(XwName);
	m_wdFt.SetNameOther(XwName);
	m_wdFt.SetColor(color);
		
	_ParagraphFormat Pgf=m_Target_Sel.GetParagraphFormat();
	//Sel.ClearFormatting();
	Pgf.SetLeftIndent(0);
	Pgf.SetRightIndent(0);
	Pgf.SetSpaceBefore(0);
	Pgf.SetSpaceBeforeAuto(0);
	Pgf.SetSpaceAfter(0);
	Pgf.SetSpaceAfterAuto(0);
	Pgf.SetLineSpacingRule(0);//1);//1.5倍行距,0单位行距
	//Pgf.SetLineSpacing(12.0);
	Pgf.SetAlignment(3);//对齐方式
	Pgf.SetWidowControl(0);
	Pgf.SetKeepWithNext(0);
	Pgf.SetKeepTogether(0);

	
	Pgf.SetPageBreakBefore(0);
	Pgf.SetNoLineNumber(0);
	long True=-1;
	Pgf.SetHyphenation(True);//支持英文换行.SetHyphenation(1);用字连接符
	Pgf.SetFirstLineIndent(0);
	Pgf.SetOutlineLevel(10);
	Pgf.SetCharacterUnitLeftIndent(0);
	
	//.CharacterUnitRightIndent = 0
	Pgf.SetCharacterUnitRightIndent(0);
	//.CharacterUnitFirstLineIndent = 0
	Pgf.SetCharacterUnitFirstLineIndent(0);
	//        .LineUnitBefore = 0
	Pgf.SetLineUnitBefore(0);
	Pgf.SetWordWrap(True);
	
	// .LineUnitAfter = 0
	Pgf.SetLineUnitAfter(0);
	//.AutoAdjustRightIndent = True
	Pgf.SetAutoAdjustRightIndent(0);
	//.DisableLineHeightGrid = False
	Pgf.SetDisableLineHeightGrid(True);
	//.FarEastLineBreakControl = True
	Pgf.SetFarEastLineBreakControl(True);
	
	//        .WordWrap = True
	Pgf.SetWordWrap(True);
	//.HangingPunctuation = True
	Pgf.SetHangingPunctuation(True);
	//.HalfWidthPunctuationOnTopOfLine = False
	Pgf.SetHalfWidthPunctuationOnTopOfLine(0);
	//.AddSpaceBetweenFarEastAndAlpha = True
	Pgf.SetAddSpaceBetweenFarEastAndAlpha(True);
	//.AddSpaceBetweenFarEastAndDigit = True
	Pgf.SetAddSpaceBetweenFarEastAndDigit(True);
	//Pgf.SetLineSpacingRule(1);//标准字距
	//.BaseLineAlignment = wdBaselineAlignAuto
	Pgf.SetBaseLineAlignment(4);
	//m_Target_Sel.SetFont(m_wdFt.DetachDispatch());

}
BOOL CyzWordOperator::GetCursorTable_Row_Col(long &TableNo, long &row, long &col)
{
	if(m_pWordApp==NULL)
		return FALSE;
	m_WordDoc.Activate();
	Selection sel=m_pWordApp->GetSelection();//GetRange();
	long startpos=sel.GetStart();
	long endpos=sel.GetEnd();
	Tables tables=m_WordDoc.GetTables();
	if(!tables.m_lpDispatch)
		return FALSE;
	long ns=tables.GetCount();
	TableNo=row=col=-1;
	for(long i=1;i<=ns;i++)
	{
		Table tab=tables.Item(i);
		Range tabrgn=tab.GetRange();
		if(tabrgn.GetStart()<=startpos && tabrgn.GetEnd()>=endpos)
		{
			Rows rows=tab.GetRows();
			Columns cols=tab.GetColumns();
			long rsize=rows.GetCount();
			long csize=cols.GetCount();
			
			for(long r=1;r<=rsize;r++)
			{
				for(long c=1;c<=csize;c++)
				{
					TRY
					{
						Cell cell=tab.Cell(r,c);
						if(!cell.m_lpDispatch)
							continue;
						Range crgn=cell.GetRange();
						if(crgn.GetStart()<=startpos && crgn.GetEnd()>=endpos)
						{
							TableNo=i;
							row=r;
							col=c;
							return TRUE;
						}
					}
					CATCH(CException, e)
					{
						continue;
					}
					END_CATCH
				}

			}
		}
	}
	TableNo=row=col=-1;
	return FALSE;
}

BOOL CyzWordOperator::GotoPageNo(int PageNo)
{
	if(m_pWordApp==NULL)
		return FALSE;
	Selection m_Target_Sel=m_pWordApp->GetSelection();
	CString PageSz;
	PageSz.Format("%d",PageNo);
	COleVariant What((short)1),Which((short)2),Count((short)1);
	COleVariant Name(PageSz),FindText(_T("^m")),Replacement(_T(""));
	Page page=m_Target_Sel.GoTo(What,Which,Count,Name);
	if(page.m_lpDispatch)
		return TRUE;
	return FALSE;
}


bool CyzWordOperator::Wordexit()
{
	m_pWordApp->m_bAutoRelease = TRUE;
	VARIANT vt;
	vt.vt = VT_ERROR;
	vt.scode = DISP_E_PARAMNOTFOUND;

	VARIANT v;
	v.vt = VT_BOOL;
	v.boolVal = VARIANT_FALSE;
	m_pWordApp->Quit(&v, &vt, &vt);
	m_pWordApp->DetachDispatch();
	m_pWordApp->ReleaseDispatch();
	m_pdispWordApp->Release();
	//COleVariant vtMissing(DISP_E_PARAMNOTFOUND, VT_ERROR); 
	//BYTE parms[] =VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
	//m_pWordApp->InvokeHelper(0x451, DISPATCH_METHOD, VT_EMPTY, NULL, parms,&vtMissing, &vtMissing, &vtMissing); 

	delete m_pWordApp;
	// TODO: 在此处添加实现代码.
	return false;
}
