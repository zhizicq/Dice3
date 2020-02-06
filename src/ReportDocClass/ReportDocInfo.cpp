// ReportDocInfo.cpp: implementation of the CReportDocInfo class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "ReportDocInfo.h"
#include "math.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CReportDocInfo::CReportDocInfo()
{
	m_PageWide=21.0;//ҳ����cm
	m_PageHigh=29.7;//ҳ��߶�cm
	m_PageLeft=3.17;//ҳ����߾�
	m_PageRight=3.17;//ҳ���ұ߾�
	m_PageTop=2.54;//ҳ�涥�߾�
	m_PageBottom=2.54;//ҳ��ױ߾�

}

CReportDocInfo::~CReportDocInfo()
{

}
//��ȡĿ¼��Ϣ
void CReportDocInfo::GetMuluInfo()
{
	if(!m_WordApp.GetWordAppObj())
	{
		AfxMessageBox("û��Word App");
			return;
	}
	//��Ŀ¼�����ӹ�ϵ
	m_WordApp.m_WordDoc.Activate();
	Hyperlinks hypers=m_WordApp.m_WordDoc.GetHyperlinks();
	long sizes=hypers.GetCount();
	if(sizes<1)
	{
		AfxMessageBox("û���Զ�Ŀ¼");
		return;
	}
	COleVariant ii((short)1);
	for(int i=1;i<=sizes;i++)
	{
		CString Info;
		ii.lVal=(short)i;
		Hyperlink hyi=hypers.Item(ii);
		CString Name=hyi.GetName();
		Range rgn=hyi.GetRange();
		Fields fields=rgn.GetFields();
		long size=fields.GetCount();
		if(size==2)
		{
			Field field=fields.Item(1);
			long sel=field.GetType();//wdFieldHyperlink 88,wdFieldPageRef = 37������
			field.GetLinkFormat();//.GetOLEFormat();
			//CString szData=field.GetData();
			CString szAdr=rgn.GetText();//hyi.GetName();//.GetAddress();
			
			Info.Format("%s\t%s",Name,szAdr);
			if(szAdr.Find("\t")>=0)
				AfxMessageBox(Info);
		}
	}
	
}
void CReportDocInfo::GetPageSet()
{
	SetCheck();
	return;
	if(!m_WordApp.GetWordAppObj())
	{
		AfxMessageBox("û��Word App");
			return;
	}
	m_WordApp.m_WordDoc.Activate();
	Sections secs=m_WordApp.m_WordDoc.GetSections();
	long sizes=secs.GetCount();
	CString txt;
	txt.Format("%d",sizes);
	AfxMessageBox(txt);
	for(long i=1;i<=sizes;i++)
	{
		Section sec=secs.Item(i);
		HeadersFooters heads=sec.GetHeaders();
		long ls=heads.GetCount();
		for(long z=1;z<=ls;z++)
		{
			HeaderFooter hf=heads.Item(z);
			BOOL isH=hf.GetIsHeader();
			if(isH)
			{
				Range rgn2=hf.GetRange();
				//rgn2.Select();
				CString ttt=rgn2.GetText();
				ttt.Replace("\n","");
				ttt.Replace("\r","");
				ttt.Replace("   ","");
				if(ttt.GetLength()>0)
					txt+=ttt+"\n";
			}
			else
				AfxMessageBox("foot");
		}
		
		//Range rgn=sec.GetRange();
		//txt=rgn.GetText();

	}
		AfxMessageBox(txt);

}
void CReportDocInfo::SetCheck()
{
	//ȫ���ַ��û�
	FindReplace("��", '(');
	FindReplace("��", ')');
	FindReplace("��","/");
	FindReplace("��","!");
	FindReplace("��","1");
	FindReplace("��","2");
	FindReplace("��","3");
	FindReplace("��","4");
	FindReplace("��","5");
	FindReplace("��","6");
	FindReplace("��","7");
	FindReplace("��","8");
	FindReplace("��","9");
	FindReplace("��","0");
	FindReplace("��"," ");
	FindReplace("��","{");
	FindReplace("��","}");
	FindReplace("^p^t", "^p",FALSE);
	FindReplace("^t^p", "^p");
	FindReplace("^p^p", "^p");
	FindReplace("^p ", "^p");
	FindReplace(" ^p", "^p");
	FindReplace("^l", "^p");//�ֶ����з�

	Check_WordShapes();//Word�Ի�ͼ�δ���
	Check_ShapePicture();//����ͼ�δ���
	Check_Tables();//��������Ϣ,���߱�

	Check_Biaoti();//����Ͷθ�ʽ����
	Check_InlineShapes();//Ƕ�������,����ʽ��,ͼ�����¶�ͬҳ
	Check_Bianhao();//��������Ŀ������
	AfxMessageBox("���");

}

void CReportDocInfo::GetParagraphs_Info()
{
	if(!m_WordApp.GetWordAppObj())
	{
		AfxMessageBox("û��Word App");
			return;
	}
	m_WordApp.m_WordDoc.Activate();
	Paragraphs Pgraphs=m_WordApp.m_WordDoc.GetParagraphs();
	int sizes=Pgraphs.GetCount();
	CString Info;
	for(long i=1;i<=sizes;i++)
	{
		Paragraph pagh=Pgraphs.Item(i);
		Range rgn=pagh.GetRange();
		Tables tabs=rgn.GetTables();
		//���Ϊ��񲻼��
		if(tabs.GetCount()>0)
			continue;
		//��ȡ���б���(��ţ���Ŀ��)
		ListParagraphs ParList=rgn.GetListParagraphs();
		long parCount=ParList.GetCount();
		if(parCount>0)//Ϊ��Ż���Ŀ��
		{
			Paragraph list=ParList.Item(1);
			Range rgn=list.GetRange();
			CString stylename=rgn.GetText();//list.GetStyleName();
			char* pch=stylename.GetBuffer(0);
			ListFormat listfor=rgn.GetListFormat();
			CString str=listfor.GetListString();//��Ŀ��
			COleVariant wdNumberParagraph((short)1);
			listfor.RemoveNumbers(wdNumberParagraph);
			stylename=str+" "+stylename;
			rgn.SetText(stylename);
			continue;
				
			long type=listfor.GetListType();//.GetListLevelNumber();
			ListTemplate listtmp=listfor.GetListTemplate();
			
			ListLevels levels=listtmp.GetListLevels();

			
			CString Name=listtmp.GetName();
			long ltab=listfor.GetListLevelNumber();
			continue;
		}
		//ȡ�����ʽ
		_ParagraphFormat Pfmt=pagh.GetFormat();
		long level=Pfmt.GetOutlineLevel();
		if(level<10)//�����⼶
		{
			continue;
		}
		_Font font=rgn.GetFont();
		CString fntName=font.GetName();
		CString fntNameAsc=font.GetNameAscii();
		float ftsize=font.GetSize();
		float size=font.GetSize();
		font.SetSize(12);
		font.SetNameAscii("Times New Roman");
		//1����0.35����=0.35278 ����=1/72*25.4
		/*
		���ţ�42,С��-36,1��-26,С1��-24,2��-22,С2��18,3��-16,С3��-15
		4��14,С4��12,5��-10.5,С5��9,6��-7.5,С6��6.5,7��5.5,8��5
		*/
		
		float l=Pfmt.GetLineSpacing();//�о�
		Pfmt.SetLineSpacingRule(1);//1.5���о�
		float lI=Pfmt.GetLeftIndent();//������
		Pfmt.SetLeftIndent(0);
		float rI=Pfmt.GetRightIndent();//������
		float fI=Pfmt.GetFirstLineIndent();//��������
		Pfmt.SetCharacterUnitFirstLineIndent(2);//��������2��

		long Ali=Pfmt.GetAlignment();//���뷽ʽ
		float fcI=Pfmt.GetCharacterUnitFirstLineIndent();//��ǰ����
		float lcI=Pfmt.GetCharacterUnitLeftIndent();//���������ַ���
		float rcI=Pfmt.GetCharacterUnitRightIndent();//���������ַ���
		float dq=Pfmt.GetSpaceBefore();
		float dh=Pfmt.GetSpaceAfter();
		CString txt,ls;
		ls=rgn.GetText();
		txt.Format("��ټ���=%d,��������%s,Ӣ������%s,�ֺ�%.1f,�о�=%.2f\n%s",level,fntName,fntNameAsc,ftsize,l,ls);
		
		//CString txt=rgn.GetText();
		//AfxMessageBox(txt);
	}
	AfxMessageBox("��������������!");

}

void CReportDocInfo::GetTableInfo()
{
	if(!m_WordApp.GetWordAppObj())
	{
		AfxMessageBox("û��Word App");
			return;
	}
	m_WordApp.m_WordDoc.Activate();
	Tables tabs=m_WordApp.m_WordDoc.GetTables();
	long tabsize=tabs.GetCount();
	if(tabsize<=0)
		return;
	for(long i=1;i<=tabsize;i++)
	{
		Table tab=tabs.Item(i);
		Range rgn=tab.GetRange();
		Cells cells=rgn.GetCells();
		Rows rows=tab.GetRows();
		rows.SetHeightRule(0);
		rows.SetHeight(0,0);
		//���е�Ԫ������
		Borders cbrds=cells.GetBorders();
		Border h=cbrds.Item(5);
		h.SetColor(RGB(0,0,0));
		h.SetLineStyle(1);//ʵ��
		h.SetLineWidth(4);//0.5P
		Border v=cbrds.Item(6);//����м�����
		v.SetLineStyle(0);
		cells.SetVerticalAlignment(1);//���е�Ԫ���������
		_Font fnt=rgn.GetFont();
		//fnt.SetColor(RGB(255,0,0));
		fnt.SetName("����");//"΢���ź�");
		fnt.SetNameAscii("Times New Roman");
		fnt.SetSize(10.5);//5��
		long wdAutoFitWindow =2;//1�������ݵ������� 2;// = 1;
		tab.AutoFitBehavior(wdAutoFitWindow);
		//�������������·
		_ParagraphFormat pfmt=rgn.GetParagraphFormat();//��ƌ��ʽ
		pfmt.SetSpaceAfterAuto(0);
		pfmt.SetSpaceAfter(0);
		pfmt.SetLineSpacingRule(0);//�����о�
		//pfmt.SetAutoAdjustRightIndent(1);//.SetAutoAdjustRightIndent(1);
		pfmt.SetDisableLineHeightGrid(0);
		//���ñ߿�
		Borders borders=tab.GetBorders();
		Border top=borders.Item(1);//����
		top.SetLineStyle(0);//�ޱ߿�
		Border left=borders.Item(2);
		left.SetLineStyle(0);//�ޱ߿�
		Border bottom=borders.Item(3);
		bottom.SetColor(RGB(0,0,0));
		bottom.SetLineStyle(1);//�еױ�
		bottom.SetLineWidth(4);//0.5P
		Border right=borders.Item(4);
		right.SetLineStyle(0);//�ޱ߿�
		Border up=borders.Item(7);//б��
		up.SetLineStyle(0);
		Border down=borders.Item(8);//б��
		down.SetLineStyle(0);
	}




}

void CReportDocInfo::GetShapesInfo()
{
	CyzWordOperator Word;
	if(!Word.GetWordAppObj())
		return;
	//��Ŀ¼
	Word.m_WordDoc.Activate();
	//Word�Ի�ͼ�Ĺ���
	Shapes sps=Word.m_WordDoc.GetShapes();
	Selection sel=Word.m_pWordApp->GetSelection();
	ShapeRange shapeRgn=sel.GetShapeRange();
	
	if(sps.m_lpDispatch==NULL)
		return ;
	long count=sps.GetCount();
	COleVariant index((short)0);
	for(long i=1;i<=count;i++)
	{
		index.iVal=(short)i;
		Shape shape=sps.Item(index);
		
		float wide=(float)(shape.GetWidth()*0.03527f);//cm
		wide=(int)((wide+0.005)*100)/100.0f;
		float high=shape.GetHeight()*0.03527f;//cm
		high=high+0.005f;
		high=(int)(100*high)/100.00f;
		long per=shape.GetZOrderPosition();
		//shape.Select(&index);
		CString Name=shape.GetName();
		/*if(Name.Find("Picture")>=0)//Group
		{
			shape.ConvertToInlineShape();
			count--;
			i--;
		}
		else */
		if(Name.Find("Group")>=0)
		{
			WrapFormat Wfmt=shape.GetWrapFormat();//.GetPictureFormat();
			long type=Wfmt.GetType();
			if(type!=4)//7Ƕ����
				Wfmt.SetType(4);//0-����wdWrapTopBottom4������
		}
	}
}
//��ȡǶ�������Ϣ
void CReportDocInfo::GetInlineShapesInfo()
{
	CyzWordOperator Word;
	if(!Word.GetWordAppObj())
		return;
	//��Ŀ¼
	Word.m_WordDoc.Activate();
	//WordǶ�����
	InlineShapes sps=Word.m_WordDoc.GetInlineShapes();
	if(sps.m_lpDispatch==NULL)
		return ;
	long count=sps.GetCount();
	for(long i=1;i<=count;i++)
	{
		InlineShape shape=sps.Item(i);
		float wide=(float)(shape.GetWidth()*0.03527f);//cm
		wide=(int)((wide+0.005)*100)/100.0f;
		float high=shape.GetHeight()*0.03527f;//cm
		high=high+0.005f;
		high=(int)(100*high)/100.00f;
		//long per=shape.GetZOrderPosition();
		//shape.Select(&index);
		//CString Name=shape.GetName();
		long type=shape.GetType();
		OLEFormat Olefmt=shape.GetOLEFormat();
		CString OleName=Olefmt.GetClassType();
		CString OleIconNmae=Olefmt.GetProgID();
	}
}

void CReportDocInfo::CheckKeywordItem()
{
	GetKeywordInfo();
	return;
	CyzWordOperator Word;
	if(!Word.GetWordAppObj())
		return;
	//��Ŀ¼
	Word.m_WordDoc.Activate();
	CString DocInfo=Word.GetDocument_CountInfo();
	//WordǶ�����
	Paragraphs pars=Word.m_WordDoc.GetParagraphs();
	if(pars.GetCount()<1)
		return;
	CString txt;
	txt.Format("%d\n%s",pars.GetCount(),DocInfo);
	AfxMessageBox(txt);


}

CString CReportDocInfo::GetKeywordInfo()
{
	CStringArray Keys,GetWords;
	Keys.Add("��Ŀ");
	Keys.Add("ѧԺ");
	Keys.Add("רҵ�༶");
	Keys.Add("ѧ������");
	Keys.Add("ѧ��");
	Keys.Add("ָ����ʦ");
	CyzWordOperator Word;
	if(!Word.GetWordAppObj())
		return "";
	//��Ŀ¼
	Word.m_WordDoc.Activate();
	CString DocInfo=Word.GetDocument_CountInfo();
	//WordǶ�����
	Paragraphs pars=Word.m_WordDoc.GetParagraphs();
	if(pars.GetCount()<1)
		return "";
	for(long i=1;i<40;i++)
	{
		Paragraph pagh=pars.Item(i);
		Range rgn=pagh.GetRange();
		CString txt=rgn.GetText();
		txt.Replace(" ","");
		//txt.Replace("\n","");
		txt.Replace("\r","");
		if(txt.GetLength()<2)
			continue;
		int pos[20],index[20];
		int c=0;
		for(int a=0;a<Keys.GetSize();a++)
		{
			int idx=txt.Find(Keys.GetAt(a));
			if(idx>=0)
			{
				index[c]=a;
				pos[c]=idx;
				c++;
			}
		}
		CString szKey;
		for(int d=0;d<c;d++)
		{
			CString ls;
			ls.Format("%s-%s",Keys.GetAt(index[d]),txt.Mid(pos[d]));
			szKey=ls+"\n";
		}
		if(szKey.GetLength())
			AfxMessageBox(szKey);
	
	}
	return "";


}

void CReportDocInfo::Check_InlineShapes()
{
	if(!m_WordApp.GetWordAppObj())
		return;
	//��Ŀ¼
	m_WordApp.m_WordDoc.Activate();
	Selection m_Target_Sel=m_WordApp.m_pWordApp->GetSelection();
	//����β
	COleVariant six((short)6),zero((short)0);
	//�����ĵ���ʼ��
	six.intVal=6;
	m_Target_Sel.HomeKey(&six,&zero);

	//WordǶ�����
	InlineShapes sps=m_WordApp.m_WordDoc.GetInlineShapes();
	if(sps.m_lpDispatch==NULL)
		return ;
	long count=sps.GetCount();
	if(count>0)
	{
		CString ls;
		ls.Format("Ƕ��ͼ����:%d\n",count);
		m_szError+=ls;
	}
	
	for(long i=1;i<=count;i++)
	{
		InlineShape shape=sps.Item(i);
		float wide=(float)(shape.GetWidth()*0.03527f);//cm
		wide=(int)((wide+0.005)*100)/100.0f;
		float high=shape.GetHeight()*0.03527f;//cm
		high=high+0.005f;
		high=(int)(100*high)/100.00f;
		if(wide>m_PageWide)
		{
			CString ls;
			ls.Format("��%dͼ����ҳ����(%.2f����)\n",i,wide);
			m_szError+=ls;
		}
		if(high>m_PageHigh)
		{
			CString ls;
			ls.Format("��%dͼ����ҳ����(%.2f)\n",i,wide);
			m_szError+=ls;
		}
		long type=shape.GetType();
		OLEFormat Olefmt=shape.GetOLEFormat();
		CString OleName=Olefmt.GetClassType();
		CString OleIconNmae=Olefmt.GetProgID();
		if(OleName.GetLength()<=0)
		{
			Range rgn=shape.GetRange();
			_ParagraphFormat Pgfmt=rgn.GetParagraphFormat();
			Pgfmt.SetSpaceBeforeAuto(FALSE);
			Pgfmt.SetSpaceAfterAuto(FALSE);
			Pgfmt.SetFarEastLineBreakControl(-TRUE);
			Pgfmt.SetAutoAdjustRightIndent(FALSE);
			Pgfmt.SetDisableLineHeightGrid(-TRUE);
			Pgfmt.SetWidowControl(FALSE);
			Pgfmt.SetLeftIndent(0);
			Pgfmt.SetRightIndent(0);
			Pgfmt.SetFirstLineIndent(0);
			Pgfmt.SetCharacterUnitFirstLineIndent(0);//.SetLeftIndent(0);
			Pgfmt.SetAlignment(1);
			Pgfmt.SetKeepWithNext(-1);//�¶����ͬҳ
		}
	}
}

void CReportDocInfo::Check_PageSet()
{
	if(!m_WordApp.GetWordAppObj())
	{
		AfxMessageBox("û��Word App");
			return;
	}
	const float cs=28.347619f;
	m_WordApp.m_WordDoc.Activate();
	PageSetup Pageset=m_WordApp.m_WordDoc.GetPageSetup();
	float width=Pageset.GetPageWidth()/cs;//��ת����
	width=((int)((width+0.005)*100))/100.0f;
	if(fabs(width-m_PageWide)>0.001)
	{
		CString ls;
		ls.Format("ҳ����Ӧ����Ϊ%.2f,��ǰ����Ϊ%.2f����\n",m_PageWide,width);
		m_szError+=ls;
	}
	float high=Pageset.GetPageHeight()/cs;
	high=((int)((high+0.005)*100))/100.0f;
	if(fabs(high-m_PageHigh)>0.001)
	{
		CString ls;
		ls.Format("ҳ��߶ȶ�Ӧ����Ϊ%.2f,��ǰ����Ϊ%.2f����\n",m_PageHigh,high);
		m_szError+=ls;

	}

	float rm=Pageset.GetRightMargin()/cs;//�ұ߾�
	rm=((int)((rm+0.005)*100))/100.0f;
	if(fabs(rm-m_PageRight)>0.001)
	{
		CString ls;
		ls.Format("ҳ���ұ߾�Ӧ����Ϊ%.2f,��ǰ����Ϊ%.2f����\n",m_PageRight,rm);
		m_szError+=ls;

	}

	float lm=Pageset.GetLeftMargin()/cs;//��߾�
	lm=((int)((lm+0.005)*100))/100.0f;
	if(fabs(lm-m_PageLeft)>0.001)
	{
		CString ls;
		ls.Format("ҳ����߾�Ӧ����Ϊ%.2f,��ǰ����Ϊ%.2f����\n",m_PageLeft,lm);
		m_szError+=ls;

	}

	float tm=Pageset.GetTopMargin()/cs;
	tm=((int)((tm+0.005)*100))/100.0f;
	if(fabs(tm-m_PageTop)>0.001)
	{
		CString ls;
		ls.Format("ҳ�涥�߾�Ӧ����Ϊ%.2f,��ǰ����Ϊ%.2f����\n",m_PageTop,tm);
		m_szError+=ls;

	}
	float bm=Pageset.GetBottomMargin()/cs;
	bm=((int)((bm+0.005)*100))/100.0f;
	if(fabs(bm-m_PageBottom)>0.001)
	{
		CString ls;
		ls.Format("ҳ��ױ߾�Ӧ����Ϊ%.2f,��ǰ����Ϊ%.2f����\n",m_PageBottom,bm);
		m_szError+=ls;

	}
	float cl=Pageset.GetCharsLine();//ÿ���ַ���
	cl=((int)((cl+0.005)*100))/100.0f;
	long l=Pageset.GetPaperSize();
}
//��������Ϣ
void CReportDocInfo::Check_Paragraphs()
{
	if(!m_WordApp.GetWordAppObj())
	{
		AfxMessageBox("û��Word App");
			return;
	}
	m_WordApp.m_WordDoc.Activate();
	Paragraphs Pgraphs=m_WordApp.m_WordDoc.GetParagraphs();
	int sizes=Pgraphs.GetCount();
	CString Info;
	int Paragrahps_Count=0;
	CString ls;
	for(long i=1;i<=sizes;i++)
	{
		Paragraph pagh=Pgraphs.Item(i);
		Range rgn=pagh.GetRange();
		Tables tabs=rgn.GetTables();
		//���Ϊ��񲻼��
		if(tabs.GetCount()>0)
			continue;
		//��ȡ���б���(��ţ���Ŀ��)
		ListParagraphs ParList=rgn.GetListParagraphs();
		long parCount=ParList.GetCount();
		if(parCount>0)//Ϊ��Ż���Ŀ��
			continue;
		//ȡ�����ʽ
		_ParagraphFormat Pfmt=pagh.GetFormat();
		long level=Pfmt.GetOutlineLevel();
		if(level<10)//�����⼶
			continue;
		Paragrahps_Count++;
		_Font font=rgn.GetFont();
		CString fntName=font.GetName();
		CString fntNameAsc=font.GetNameAscii();
		if(fntName.CompareNoCase("����")!=0)
		{
			ls.Format("��%d�����������岻�� ����,����%s\n",i,fntName);
			this->m_szError+=ls;
		}
		if(fntNameAsc.CompareNoCase("Times New Roman")!=0)
		{
			ls.Format("��%d����Ӣ�����岻 Times New Roman ,����%s\n",i,fntNameAsc);
			m_szError+=ls;
		}
		float size=font.GetSize();
		if(size!=12)
		{
			ls.Format("��%d��������ߴ粻��С4��(12.0),����%.1f\n",i,size);
			this->m_szError+=ls;
		}
		//1����0.35����=0.35278 ����=1/72*25.4
		/*
		���ţ�42,С��-36,1��-26,С1��-24,2��-22,С2��18,3��-16,С3��-15
		4��14,С4��12,5��-10.5,С5��9,6��-7.5,С6��6.5,7��5.5,8��5
		*/
		
		float l=Pfmt.GetLineSpacing();//�о�
		long rhj=Pfmt.GetLineSpacingRule();
		if(rhj!=1)
		{
			ls.Format("��%d�����о಻�ǹ����о���(1.5��),����%d\n",i,rhj);
			this->m_szError+=ls;
		}

		float lI=Pfmt.GetLeftIndent();//������
		if(lI>0.0)
		{
			ls.Format("��%d�������������(%.2f)\n",i,lI);
			this->m_szError+=ls;

		}

		Pfmt.SetLeftIndent(0);
		float rI=Pfmt.GetRightIndent();//������
		if(rI>0.0)
		{
			ls.Format("��%d�������������(%.2f)\n",i,rI);
			this->m_szError+=ls;

		}

		float fI=Pfmt.GetFirstLineIndent();//��������
		fI=Pfmt.GetCharacterUnitFirstLineIndent();
		if(fI!=2.0)
		{
			ls.Format("��%d����������������2���ַ�\n",i);
			m_szError+=ls;
		}
		//CString txt=rgn.GetText();

		//Pfmt.SetCharacterUnitFirstLineIndent(2);//��������2��

		long Ali=Pfmt.GetAlignment();//���뷽ʽ
		ls.Format("��%d������뷽ʽ,%d\n",i,Ali);
		m_szError+=ls;

	

		float fcI=Pfmt.GetCharacterUnitFirstLineIndent();//��ǰ����
		float lcI=Pfmt.GetCharacterUnitLeftIndent();//���������ַ���
		float rcI=Pfmt.GetCharacterUnitRightIndent();//���������ַ���
		float dq=Pfmt.GetSpaceBefore();
		float dh=Pfmt.GetSpaceAfter();
		
		//AfxMessageBox(txt);
	}
	AfxMessageBox("��������������!");
}

void CReportDocInfo::FindReplace(CString Keyword, CString Space,BOOL bWhile)
{
	if(!m_WordApp.GetWordAppObj())
	{
		AfxMessageBox("û��Word App");
			return;
	}
	m_WordApp.m_WordDoc.Activate();
	
	Selection m_Target_Sel=m_WordApp.m_pWordApp->GetSelection();
	//����β
	COleVariant six((short)6),zero((short)0);
	//�����ĵ���ʼ��
	six.intVal=6;
	m_Target_Sel.HomeKey(&six,&zero);
	m_Target_Sel=m_WordApp.m_pWordApp->GetSelection();
	//��ͷ��ʼ
	BOOL bt=false;
	int cs=0;
	do
	{
		Find w_Find=m_Target_Sel.GetFind();
		w_Find.ClearFormatting();
		//Keyword="^p";//���η�
		COleVariant FindText(Keyword);
		COleVariant MatchCase((short)1), MatchWholeWord((short)0), MatchWildcards((short)0);
		COleVariant MatchSoundsLike((short)0), MatchAllWordForms((short)0),Forward((short)1),Wrap((short)1);////����ת
		COleVariant Format((short)0), MatchKashida((short)0),MatchDiacritics((short)0);
		COleVariant MatchAlefHamza((short)0), MatchControl((short)0);
		COleVariant Replace((short)2);//2-ȫ���滻��1-�滻һ����0���滻
		COleVariant ReplaceWith(Space);
		COleVariant Direction((short)1);//������
		
		int ns=0;
		
		bt=w_Find.Execute(FindText,MatchCase, MatchWholeWord, MatchWildcards,MatchSoundsLike, 
			MatchAllWordForms,Forward,Wrap,Format,ReplaceWith, Replace, MatchKashida,MatchDiacritics,MatchAlefHamza,
			MatchControl);
		cs++;
	}while(bt && cs<10 && bWhile);
}

void CReportDocInfo::Check_Biaoti()
{
	if(!m_WordApp.GetWordAppObj())
	{
		AfxMessageBox("û��Word App");
			return;
	}
	m_WordApp.m_WordDoc.Activate();
	Paragraphs Pgraphs=m_WordApp.m_WordDoc.GetParagraphs();
	int sizes=Pgraphs.GetCount();
	CString Info;
	int Paragrahps_Count=0;
	CString ls;
	int count=0;
	int levels[10]={0};
	BOOL bCheckBiaoti=FALSE;
	for(long i=1;i<=sizes;i++)
	{
		Paragraph pagh=Pgraphs.Item(i);
		Range rgn=pagh.GetRange();
		InlineShapes shapes=rgn.GetInlineShapes();
		if(shapes.GetCount()>0)
		{
			//rgn.Select();
			Characters chars=rgn.GetCharacters();
			int ns=chars.GetCount();
			if(ns<3)
				continue;
			/*
				
			InlineShape shape=shapes.Item(1);
			OLEFormat Olefmt=shape.GetOLEFormat();
			CString OleName=Olefmt.GetClassType();
			if(OleName.GetLength()<=0)
				continue;
			*/
		}
		Tables tabs=rgn.GetTables();
		//���Ϊ��񲻼��
		if(tabs.GetCount()>0)
			continue;
		//��ȡ���б���(��ţ���Ŀ��)
		ListParagraphs ParList=rgn.GetListParagraphs();
		long parCount=ParList.GetCount();
		if(parCount>0)//Ϊ��Ż���Ŀ��
			continue;
		//ȡ�����ʽ
		_ParagraphFormat Pfmt=pagh.GetFormat();
		_Font font=rgn.GetFont();
		float FLIChars=0;
		long level=pagh.GetOutlineLevel();
		if(level<10)
		{
			bCheckBiaoti=TRUE;//��⵽����
			levels[level]++;
			count++;
			float fntsize=18;
			float qhk=12.0;//ǰ���
			BOOL PageBreakBefore=0;
			switch(level)
			{
			case 1:
				fntsize=18;
				qhk=17.0;
				FLIChars=0;
				PageBreakBefore=-1;
				break;
			case 2:
				fntsize=15;
				qhk=14;
				FLIChars=0;
				break;
			case 3:
				fntsize=14;
				FLIChars=2;
				break;
			default:
				fntsize=12;
				break;
			}
		
			font.SetName("����");
			font.SetSize(fntsize);
			font.SetNameAscii("Times New Roman");
			font.SetBold(0);
			//1����0.35����=0.35278 ����=1/72*25.4
			/*
			���ţ�42,С��-36,1��-26,С1��-24,2��-22,С2��18,3��-16,С3��-15
			4��14,С4��12,5��-10.5,С5��9,6��-7.5,С6��6.5,7��5.5,8��5
			*/
			if(level>1 && level<10)
			{
				Pfmt.SetLeftIndent(0);
				Pfmt.SetRightIndent(0);
				Pfmt.SetSpaceBeforeAuto(FALSE);
				Pfmt.SetSpaceAfterAuto(FALSE);
				Pfmt.SetFarEastLineBreakControl(-TRUE);
				Pfmt.SetFirstLineIndent(0);
				Pfmt.SetCharacterUnitFirstLineIndent(FLIChars);//.SetLeftIndent(0);
			}
			Pfmt.SetSpaceBefore(qhk);
			Pfmt.SetSpaceBeforeAuto(0);
			Pfmt.SetSpaceAfter(qhk);
			Pfmt.SetSpaceAfterAuto(0);
			Pfmt.SetWidowControl(0);
			Pfmt.SetLineSpacingRule(5);
			Pfmt.SetPageBreakBefore(PageBreakBefore);
			Pfmt.SetKeepWithNext(-1);
			Pfmt.SetKeepTogether(-1);
			continue;
		}
		else if(bCheckBiaoti)//��������
		{
			_Font font=rgn.GetFont();
			CString fntName=font.GetName();
			CString fntNameAsc=font.GetNameAscii();
			font.SetName("����");
			font.SetSize(12);
			font.SetNameAscii("Times New Roman");
			//1����0.35����=0.35278 ����=1/72*25.4
			/*
			���ţ�42,С��-36,1��-26,С1��-24,2��-22,С2��18,3��-16,С3��-15
			4��14,С4��12,5��-10.5,С5��9,6��-7.5,С6��6.5,7��5.5,8��5
			*/
			// .LeftIndent = CentimetersToPoints(0)
			Pfmt.SetLeftIndent(0);
			//.RightIndent = CentimetersToPoints(0)
			Pfmt.SetRightIndent(0);
			//.SpaceBefore = 0
			Pfmt.SetSpaceBefore(0);
			//.SpaceBeforeAuto = False
			Pfmt.SetSpaceBeforeAuto(FALSE);
			//.SpaceAfter = 0
			Pfmt.SetSpaceAfter(0);
			//.SpaceAfterAuto = False
			Pfmt.SetSpaceAfterAuto(FALSE);
			//.LineSpacingRule = wdLineSpace1pt5
			Pfmt.SetLineSpacingRule(1);//1.5
        //.Alignment = wdAlignParagraphJustify
        //.WidowControl = False
			Pfmt.SetWidowControl(FALSE);
        //.KeepWithNext = False
			Pfmt.SetKeepWithNext(FALSE);
        //.KeepTogether = False
			Pfmt.SetKeepTogether(FALSE);
        //.PageBreakBefore = False
			Pfmt.SetPageBreakBefore(FALSE);
        //.NoLineNumber = False
			Pfmt.SetNoLineNumber(FALSE);
        //.Hyphenation = True
			Pfmt.SetHyphenation(-TRUE);
        //.FirstLineIndent = CentimetersToPoints(0.35)
			Pfmt.SetFirstLineIndent(0);
        //.OutlineLevel = wdOutlineLevelBodyText
       //.CharacterUnitLeftIndent = 0
			Pfmt.SetCharacterUnitLeftIndent(0);
        //.CharacterUnitRightIndent = 0
			Pfmt.SetCharacterUnitRightIndent(0);
       //.CharacterUnitFirstLineIndent = 2
			Pfmt.SetCharacterUnitFirstLineIndent(2);
       // .LineUnitBefore = 0
			Pfmt.SetLeftIndent(0);
       // .LineUnitAfter = 0
			Pfmt.SetLineUnitAfter(0);
      //  .AutoAdjustRightIndent = False
			Pfmt.SetAutoAdjustRightIndent(FALSE);
       // .DisableLineHeightGrid = True
			Pfmt.SetDisableLineHeightGrid(-TRUE);
       //.FarEastLineBreakControl = True
			Pfmt.SetFarEastLineBreakControl(-TRUE);
       // .WordWrap = True
			Pfmt.SetWordWrap(-TRUE);
       //.HangingPunctuation = True
			Pfmt.SetHangingPunctuation(-TRUE);
      //  .HalfWidthPunctuationOnTopOfLine = False
			Pfmt.SetHalfWidthPunctuationOnTopOfLine(FALSE);
      //  .AddSpaceBetweenFarEastAndAlpha = True
			Pfmt.SetAddSpaceBetweenFarEastAndAlpha(-TRUE);
      //  .AddSpaceBetweenFarEastAndDigit = True
			Pfmt.SetAddSpaceBetweenFarEastAndDigit(-TRUE);
      //  .BaseLineAlignment = wdBaselineAlignAuto
			Pfmt.SetBaseLineAlignment(4);
	//Pfmt.SetCharacterUnitLeftIndent(0);
			//Pfmt.SetCharacterUnitRightIndent(0);
			//float l=Pfmt.GetLineSpacing();//�о�
	//Pfmt.SetLineSpacingRule(1);//1.5���о�
			//float lI=Pfmt.GetLeftIndent();//������
			//float rI=Pfmt.GetRightIndent();//������
			//float fI=Pfmt.GetFirstLineIndent();//��������
	//Pfmt.SetCharacterUnitFirstLineIndent(2);//��������2��
	//Pfmt.SetDisableLineHeightGrid(FALSE);//-TRUE);
			
			if(Pfmt.GetAlignment()!=1)
				Pfmt.SetAlignment(3);//���˶���wdAlignParagraphJustify = 3
		}

	}
	ls.Format("������=%d",count);
//	AfxMessageBox(ls);
}
//���Word�Ի�ͼ��
void CReportDocInfo::Check_WordShapes()
{
	if(!m_WordApp.GetWordAppObj())
		return;
	//��Ŀ¼
	m_WordApp.m_WordDoc.Activate();
	//Word�Ի�ͼ�Ĺ���
	Shapes sps=m_WordApp.m_WordDoc.GetShapes();
	if(sps.m_lpDispatch==NULL)
		return ;
	long count=sps.GetCount();
	COleVariant index((short)0);
	int Others=0;
	for(long i=1;i<=count;i++)
	{
		index.iVal=(short)i;
		Shape shape=sps.Item(index);
		
		float wide=(float)(shape.GetWidth()*0.03527f);//cm
		wide=(int)((wide+0.005)*100)/100.0f;
		float high=shape.GetHeight()*0.03527f;//cm
		high=high+0.005f;
		high=(int)(100*high)/100.00f;
		long per=shape.GetZOrderPosition();
		float top=shape.GetTop();
		//shape.Select(&index);
		CString Name=shape.GetName();
		if(Name.Find("Picture")>=0)//Group
			continue;//������ͼ�β�����
		else if(Name.Find("Group")>=0)
		{
			WrapFormat Wfmt=shape.GetWrapFormat();//.GetPictureFormat();
			long type=Wfmt.GetType();
			if(type!=4)//7Ƕ����
			{
				Wfmt.SetType(4);//0-����wdWrapTopBottom4������
				shape.SetRelativeVerticalPosition(2);//��Զ���
				shape.SetTop(0);
			}
		}
		else
			Others++;
	}
	if(Others>0)
	{
		CString txt;
		txt.Format("����Word���Ի�ͼ��%d��,������Ҫ��ɾ��!\n",Others);
		AfxMessageBox(txt);
	}
}
//����ͼ��תΪǶ����
void CReportDocInfo::Check_ShapePicture()
{
	if(!m_WordApp.GetWordAppObj())
		return;
	//��Ŀ¼
	m_WordApp.m_WordDoc.Activate();
	//Word�Ի�ͼ�Ĺ���
	Shapes sps=m_WordApp.m_WordDoc.GetShapes();
	if(sps.m_lpDispatch==NULL)
		return ;
	long count=sps.GetCount();
	COleVariant index((short)0);
	for(long i=1;i<=count;i++)
	{
		index.iVal=(short)i;
		Shape shape=sps.Item(index);
		
		float wide=(float)(shape.GetWidth()*0.03527f);//cm
		wide=(int)((wide+0.005)*100)/100.0f;
		float high=shape.GetHeight()*0.03527f;//cm
		high=high+0.005f;
		high=(int)(100*high)/100.00f;
		long per=shape.GetZOrderPosition();
		//shape.Select(&index);
		CString Name=shape.GetName();
		if(Name.Find("Picture")>=0)//Group
		{
			shape.ConvertToInlineShape();
			i--;
			count--;
		}
	}
}

void CReportDocInfo::Check_Tables()
{
	if(!m_WordApp.GetWordAppObj())
	{
		AfxMessageBox("û��Word App");
			return;
	}
	m_WordApp.m_WordDoc.Activate();
	Tables tabs=m_WordApp.m_WordDoc.GetTables();
	long tabsize=tabs.GetCount();
	if(tabsize<=0)
		return;
	for(long i=1;i<=tabsize;i++)
	{
		Table tab=tabs.Item(i);
		//tab.Select();
		Range rgn=tab.GetRange();
		Cells cells=rgn.GetCells();
		Rows rows=tab.GetRows();
		rows.SetAlignment(1);//���о���
		rows.SetHeightRule(0);
		rows.SetHeight(0,0);
		//���е�Ԫ������
		Borders cbrds=cells.GetBorders();
		
		if(rows.GetCount()>1)
		{
			Border h=cbrds.Item(5);
			h.SetColor(RGB(0,0,0));
			h.SetLineStyle(1);//ʵ��
			long w=h.GetLineWidth();
			h.SetLineWidth(4);//4,8,12,18 0.5P
		}
		
		Border v=cbrds.Item(6);//����м�����
		v.SetLineStyle(0);
		cells.SetVerticalAlignment(1);//���е�Ԫ���������
		_Font fnt=rgn.GetFont();
		//fnt.SetColor(RGB(255,0,0));
		fnt.SetName("����");//"΢���ź�");
		fnt.SetNameAscii("Times New Roman");
		fnt.SetSize(10.5);//5��
		long wdAutoFitWindow =1;//1�������ݵ������� 2;// = 1;
		tab.AutoFitBehavior(wdAutoFitWindow);
		//�������������·
		_ParagraphFormat pfmt=rgn.GetParagraphFormat();//��ƌ��ʽ
		pfmt.SetSpaceAfterAuto(0);
		pfmt.SetSpaceAfter(0);
		pfmt.SetLineSpacingRule(0);//�����о�
		//pfmt.SetAutoAdjustRightIndent(1);//.SetAutoAdjustRightIndent(1);
		pfmt.SetDisableLineHeightGrid(0);
		//���ñ߿�
		Borders borders=tab.GetBorders();
		Border top=borders.Item(1);//����
		top.SetLineStyle(0);//�ޱ߿�
		Border left=borders.Item(2);
		left.SetLineStyle(0);//�ޱ߿�
		Border bottom=borders.Item(3);
		bottom.SetColor(RGB(0,0,0));
		bottom.SetLineStyle(1);//�еױ�
		bottom.SetLineWidth(4);//0.5P
		Border right=borders.Item(4);
		right.SetLineStyle(0);//�ޱ߿�
		Border up=borders.Item(7);//б��
		up.SetLineStyle(0);
		Border down=borders.Item(8);//б��
		down.SetLineStyle(0);
	}
}

void CReportDocInfo::Check_Bianhao()
{
	if(!m_WordApp.GetWordAppObj())
	{
		AfxMessageBox("û��Word App");
			return;
	}
			//1����0.35����=0.35278 ����=1/72*25.4
			/*
			���ţ�42,С��-36,1��-26,С1��-24,2��-22,С2��18,3��-16,С3��-15
			4��14,С4��12,5��-10.5,С5��9,6��-7.5,С6��6.5,7��5.5,8��5
			*/
	COleVariant Num((short)1);
	m_WordApp.m_WordDoc.Activate();
	ListParagraphs ListParas=m_WordApp.m_WordDoc.GetListParagraphs();
	long sizes=ListParas.GetCount();
	float dp=24,tabp=24;
	CString format;
	for(long i=1;i<=sizes;i++)
	{
		Paragraph ListPar=ListParas.Item(i);
		
		//Paragraph list=ParList.Item(1);
		Range rgn=ListPar.GetRange();
		//rgn.Select();
		CString stylename=rgn.GetText();//list.GetStyleName();
		ListFormat listfor=rgn.GetListFormat();

		//format=listfor.GetListString();
		long type=listfor.GetListType();//.GetListLevelNumber();
		//listfor.RemoveNumbers(&Num);
		ListTemplate listtmp=listfor.GetListTemplate();
		ListLevels levels=listtmp.GetListLevels();
		ListLevel lev=levels.Item(1);
		//�޸�
		format=lev.GetNumberFormat();
		if(format.Find("?")<0)
		{
			format.Replace("��","(");
			format.Replace("��",")");
			lev.SetNumberFormat(format);
			long style=lev.GetNumberStyle();
			lev.SetNumberStyle(style);
			lev.SetStartAt(1);
		}
		lev.SetNumberPosition(dp);
		lev.SetAlignment(0);
		lev.SetTextPosition(dp+tabp);
		lev.SetTabPosition(tabp);
		lev.SetResetOnHigher(0);	
	}
}

void CReportDocInfo::Check_TabCharPos()
{
	if(!m_WordApp.GetWordAppObj())
	{
		AfxMessageBox("û��Word App");
			return;
	}
	int nS,nEd;
	
	m_WordApp.GetStartEndPos(nS,nEd);
	CString txt=m_WordApp.GetStartAndEndText(nS,nEd);
	txt=m_WordApp.Get_All_Text();
	nS=1;
	int size=txt.GetLength();
	long pos=0;
	for(int i=0;i<size;i++)
	{
		if(txt.GetAt(i)=='A')
			break;
		else if(txt.GetAt(i)==7)
			continue;
		else 
		if(txt.GetAt(i)>=0)
		{
			pos++;
		}
		else
		{
			pos++,i++;
		}
	}
	m_WordApp.SetSelectPos(nS+pos-1,nS+pos-1);
	return;
	if(txt.Find("\007")>=0)
	{
		char buf[400];
		sprintf(buf,"%s",txt);
		AfxMessageBox(txt);
	}
	Selection sel=m_WordApp.m_pWordApp->GetSelection();
	nS=sel.GetStart();


}
