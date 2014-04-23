/* NOTE: You will need to add the -GX flag to your Project Options in Project->Settings->C/C++*/

// object.cpp : Implementation of Cobject
#include "stdafx.h"
#include "Fixture.h"
#include "object.h"
#include <string>
#include <map>
#include "Part.h"
#include "FixtureDLG.h"//添加对话框的头文件
/////////////////////////////////////////////////////////////////////////////
// Cobject

void Cobject::AddUserInterface()
{
	AddToolbar();
	AddMenus();
}

void Cobject::RemoveUserInterface()
{
	RemoveMenus();
	RemoveToolbar();
}

void Cobject::AddMenus()
{
	long retval = 0;
	VARIANT_BOOL ok;
	long type;
	long position;
	CComBSTR menu;
	CComBSTR method;
	CComBSTR update;
	CComBSTR hint;

	// Add menu for main frame
	type = swDocNONE;
	position = 3;
	menu.LoadString(IDS_FIXTURE_MENU);
	m_iSldWorks->AddMenu(type, menu, position, &retval);

	position = -1;
	menu.LoadString(IDS_FIXTURE_START_NOTEPAD_ITEM);
	method.LoadString(IDS_FIXTURE_START_NOTEPAD_METHOD);
	hint.LoadString(IDS_FIXTURE_START_NOTEPAD_HINT);
	m_iSldWorks->AddMenuItem2(type, m_swCookie, menu, position, method, update, hint, &ok);

	// Add menu for part frame
	type = swDocPART;
	position = 5;
	menu.LoadString(IDS_FIXTURE_MENU);
	m_iSldWorks->AddMenu(type, menu, position, &retval);

	position = -1;
	menu.LoadString(IDS_FIXTURE_START_NOTEPAD_ITEM);
	method.LoadString(IDS_FIXTURE_START_NOTEPAD_METHOD);
	hint.LoadString(IDS_FIXTURE_START_NOTEPAD_HINT);
	m_iSldWorks->AddMenuItem2(type, m_swCookie, menu, position, method, update, hint, &ok);

	//在“零件”窗口新增加“机床夹具”菜单，并生成“零件设计”子菜单。
	position = -1;
	menu.LoadString(IDS_FIXTURE_START_ITEM);
	method.LoadString(IDS_FIXTURE_START_METHOD);
	hint.LoadString(IDS_FIXTURE_START_HINT);
	m_iSldWorks->AddMenuItem2(type, m_swCookie, menu, position, method, update, hint, &ok);

	position = -1;
	menu.LoadString(IDS_FIXTURE_YAKUAI_ITEM);
	method.LoadString(IDS_FIXTURE_YAKUAI_METHOD);
	hint.LoadString(IDS_FIXTURE_YAKUAI_HINT);
	m_iSldWorks->AddMenuItem2(type, m_swCookie, menu, position, method, update, hint, &ok);

		position = -1;
	menu.LoadString(IDS_FIXTURE_PXLUN_ITEM);
	method.LoadString(IDS_FIXTURE_PXLUN_METHOD);
	hint.LoadString(IDS_FIXTURE_PXLUN_HINT);
	m_iSldWorks->AddMenuItem2(type, m_swCookie, menu, position, method, update, hint, &ok);


		position = -1;
	menu.LoadString(IDS_FIXTURE_DWJIAN_ITEM);
	method.LoadString(IDS_FIXTURE_DWJIAN_METHOD);
	hint.LoadString(IDS_FIXTURE_DWJIAN_HINT);
	m_iSldWorks->AddMenuItem2(type, m_swCookie, menu, position, method, update, hint, &ok);


		position = -1;
	menu.LoadString(IDS_FIXTURE_DXJIAN_ITEM);
	method.LoadString(IDS_FIXTURE_DXJIAN_METHOD);
	hint.LoadString(IDS_FIXTURE_DXJIAN_HINT);
	m_iSldWorks->AddMenuItem2(type, m_swCookie, menu, position, method, update, hint, &ok);


	

	// Add menu for assembly frame
	type = swDocASSEMBLY;
	position = 5;
	menu.LoadString(IDS_FIXTURE_MENU);
	m_iSldWorks->AddMenu(type, menu, position, &retval);

	position = -1;
	menu.LoadString(IDS_FIXTURE_START_NOTEPAD_ITEM);
	method.LoadString(IDS_FIXTURE_START_NOTEPAD_METHOD);
	hint.LoadString(IDS_FIXTURE_START_NOTEPAD_HINT);
	m_iSldWorks->AddMenuItem2(type, m_swCookie, menu, position, method, update, hint, &ok);

	// Add menu for drawing frame
	type = swDocDRAWING;
	position = 5;
	menu.LoadString(IDS_FIXTURE_MENU);
	m_iSldWorks->AddMenu(type, menu, position, &retval);
	position = -1;
	menu.LoadString(IDS_FIXTURE_START_NOTEPAD_ITEM);
	method.LoadString(IDS_FIXTURE_START_NOTEPAD_METHOD);
	hint.LoadString(IDS_FIXTURE_START_NOTEPAD_HINT);
	m_iSldWorks->AddMenuItem2(type, m_swCookie, menu, position, method, update, hint, &ok);

}

void Cobject::RemoveMenus()
{
	VARIANT_BOOL ok;
	CComBSTR menu;

	menu.LoadString(IDS_FIXTURE_MENU);
	m_iSldWorks->RemoveMenu(swDocNONE, menu, NULL, &ok);
	m_iSldWorks->RemoveMenu(swDocPART, menu, NULL, &ok);
	m_iSldWorks->RemoveMenu(swDocASSEMBLY, menu, NULL, &ok);
	m_iSldWorks->RemoveMenu(swDocDRAWING, menu, NULL, &ok);
}

void Cobject::AddToolbar()
{
	VARIANT_BOOL stat = VARIANT_FALSE;
	VARIANT_BOOL ok;
	HRESULT hres;

	CComBSTR title;
	CComBSTR callback;
	CComBSTR update;
	CComBSTR tip;
	CComBSTR hint;
	if (m_lToolbarID == 0)
	{
		title.LoadString(IDS_FIXTURE_TOOLBAR_TITLE);
		callback.LoadString(IDS_FIXTURE_TOOLBAR_CALLBACK);
		update.LoadString(IDS_FIXTURE_TOOLBAR_UPDATE);
		tip.LoadString(IDS_FIXTURE_TOOLBAR_TIP);
		hint.LoadString(IDS_FIXTURE_TOOLBAR_HINT);

		hres = m_iSldWorks->AddToolbar3(m_swCookie, title, IDR_TOOLBAR_SMALL, IDR_TOOLBAR_LARGE, -1,
										swDocTemplateTypeNONE|swDocTemplateTypePART|swDocTemplateTypeASSEMBLY|swDocTemplateTypeDRAWING, &m_lToolbarID);

		m_iSldWorks->AddToolbarCommand2(m_swCookie, m_lToolbarID, 0, callback, update, tip, hint, &ok);
		m_iSldWorks->AddToolbarCommand2(m_swCookie, m_lToolbarID, 1, callback, update, tip, hint, &ok);
		m_iSldWorks->AddToolbarCommand2(m_swCookie, m_lToolbarID, 2, callback, update, tip, hint, &ok);

	}
}

void Cobject::RemoveToolbar()
{
	if (m_lToolbarID != 0)
	{
		VARIANT_BOOL stat;

		HRESULT hres = m_iSldWorks->RemoveToolbar2( m_swCookie, m_lToolbarID, &stat);
		m_lToolbarID = 0;
	}
}

void Cobject::AttachEventHandlers()
{
	this->m_libid = LIBID_SldWorks;
	this->m_wMajorVerNum = GetSldWorksTlbMajor();
	this->m_wMinorVerNum = 0;

	// Connect to the SldWorks event sink
	DispEventAdvise(m_iSldWorks);

	// Connect event handlers to all previously open documents
	TMapIUnknownToDocument::iterator iter;

	CComPtr<IModelDoc2> iModelDoc2;
	m_iSldWorks->IGetFirstDocument2(&iModelDoc2);
	while (iModelDoc2 != NULL)
	{
		iter = m_swDocMap.find(iModelDoc2);
		if (iter == m_swDocMap.end())
		{
			long docType = 0;
			iModelDoc2->GetType(&docType);
			switch (docType)
			{
			case swDocPART:
				{
					// Add this to the map, if it is a Part
					CComQIPtr<IPartDoc, &__uuidof(IPartDoc)> iDoc(iModelDoc2);
					if (iDoc)
					{
						CComObject<CSwPart> *pDoc;
						CComObject<CSwPart>::CreateInstance( &pDoc);
						pDoc->Init((CComObject<Cobject>*)this, iModelDoc2);
						m_swDocMap.insert(TMapIUnknownToDocument::value_type(iModelDoc2, pDoc));
						iModelDoc2.p->AddRef();
						pDoc->AddRef();
					}
				}
				break;
			}
		}

		CComPtr<IModelDoc2> iNextModelDoc2;
		iModelDoc2->IGetNext(&iNextModelDoc2);
		iModelDoc2 = iNextModelDoc2;
	}
}

void Cobject::DetachEventHandlers()
{
	// Disconnect from the SldWorks event sink
	DispEventUnadvise(m_iSldWorks);

	// Disconnect all event handlers
	TMapIUnknownToDocument::iterator iter;

	iter = m_swDocMap.begin();

	for (iter = m_swDocMap.begin(); iter != m_swDocMap.end(); /*iter++*/) //The iteration is performed inside the loop
	{
		TMapIUnknownToDocument::value_type obj = *iter;
		iter++;
		switch (obj.second->GetType())
		{
		case swDocPART:
			{
				CComObject<CSwPart> *pDoc = (CComObject<CSwPart>*)obj.second;
				pDoc->DetachEventHandlers();
			}
			break;
		}
	}
}

void Cobject::DetachEventHandler(IUnknown *iUnk)
{
	TMapIUnknownToDocument::iterator iter;

	iter = m_swDocMap.find(iUnk);
	if (iter != m_swDocMap.end())
	{
		TMapIUnknownToDocument::value_type obj = *iter;
		obj.first->Release();
		switch (obj.second->GetType())
		{
		case swDocPART:
			{
				CComObject<CSwPart> *pDoc = (CComObject<CSwPart>*)obj.second;
				pDoc->Release();
			}
			break;
		}
		m_swDocMap.erase(iter);
	}
}

/////////////////////////////////////////////////////////////////////////////
// ISwAddin implementation

STDMETHODIMP Cobject::ConnectToSW(IDispatch * ThisSW, LONG Cookie, VARIANT_BOOL * IsConnected)
{
	VARIANT_BOOL status;

	if (IsConnected == NULL)
		return E_POINTER;

	
	m_swCookie = Cookie;
	m_iSldWorks = CComQIPtr<ISldWorks, &__uuidof(ISldWorks)>(ThisSW);
	if (m_iSldWorks != NULL)
	{	
		m_iSldWorks->SetAddinCallbackInfo((long)_Module.GetModuleInstance(), static_cast<Iobject*>(this), m_swCookie, &status);
		USES_CONVERSION;
		CComBSTR bstrNum;
		std::string strNum;
		char *buffer;

		m_iSldWorks->RevisionNumber(&bstrNum);

		strNum = W2A(bstrNum);
		m_swMajNum = strtol(strNum.c_str(), &buffer, ID_SLDWORKS_TLB_MAJOR );

		strNum = buffer;
		strNum = strNum.substr(1);
		m_swMinNum = strtol(strNum.c_str(), &buffer, ID_SLDWORKS_TLB_MINOR );

		
		AttachEventHandlers();
		AddUserInterface();
	}

	return S_OK;
}

STDMETHODIMP Cobject::DisconnectFromSW(VARIANT_BOOL * IsDisconnected)
{
	if (IsDisconnected == NULL)
		return E_POINTER;

	if (m_iSldWorks != NULL)
	{
		DetachEventHandlers();
		RemoveUserInterface();
		m_iSldWorks.Release();
	}

	return E_NOTIMPL;
}


/////////////////////////////////////////////////////////////////////////////
// DSldWorksEvents implementation
STDMETHODIMP Cobject::DocumentLoadNotify(BSTR docTitle, BSTR docPath)
{
	USES_CONVERSION;

	CComBSTR bstrTitle(docTitle);
	std::string strTitle = W2A(docTitle);

	TMapIUnknownToDocument::iterator iter;

	CComPtr<IModelDoc2> iModelDoc2;
	m_iSldWorks->IGetFirstDocument2(&iModelDoc2);
	while (iModelDoc2 != NULL)
	{
		iModelDoc2->GetTitle(&bstrTitle);
		if (strTitle == W2A(bstrTitle))
		{
			iter = m_swDocMap.find(iModelDoc2);
			if (iter == m_swDocMap.end())
			{
				ATLTRACE("\tCobject::DocumentLoadNotify current part not found\n");
				long docType = 0;
				iModelDoc2->GetType(&docType);
				switch (docType)
				{
				case swDocPART:
					{
						// Add this to the map, if it is a Part
						CComQIPtr<IPartDoc, &__uuidof(IPartDoc)> iDoc(iModelDoc2);
						if (iDoc)
						{
							CComObject<CSwPart> *pDoc;
							CComObject<CSwPart>::CreateInstance( &pDoc);
							pDoc->Init((CComObject<Cobject>*)this, iModelDoc2);
							m_swDocMap.insert(TMapIUnknownToDocument::value_type(iModelDoc2, pDoc));
							iModelDoc2.p->AddRef();
							pDoc->AddRef();
						}
					}
					break;
				}
			}
		}

		CComPtr<IModelDoc2> iNextModelDoc2;
		iModelDoc2->IGetNext(&iNextModelDoc2);
		iModelDoc2 = iNextModelDoc2;
	}

	ATLTRACE("\tCobject::DocumentLoadNotify called\n");
	return 0;
}


/////////////////////////////////////////////////////////////////////////////
// Iobject implementation

STDMETHODIMP Cobject::StartNotepad()
{
	// TODO: Add your implementation code here
	::WinExec("Notepad.exe", SW_SHOW);

	return S_OK;
}

STDMETHODIMP Cobject::ToolbarUpdate(long *status)
{
	HRESULT retval = S_OK;

	*status = 1;

	return S_OK;
}

STDMETHODIMP Cobject::Part_design()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here
	FixtureDLG dlgx;//实例化对话框的类
	dlgx.DoModal();//显示模态对话框

	if (dlgx.Drawing == TRUE)
	{


	CComPtr<IModelDoc2>pModelDoc;   //文档接口

	HRESULT hres;
	long error;
	long option=swOpenDocOptions_Silent;
	long warning;
	long retval;
	short retval_rebuild;

	CComPtr<IDimension>m_Dimension;
	CComPtr<IPartDoc>pPartDoc;//定义pPartDoc指针
	CComBSTR filename(_T("e:\\bd05.SLDPRT")); //模板文件存放地址

	m_iSldWorks->OpenDoc6(filename,swDocPART,option,NULL,&error,&warning,&pModelDoc);

	//hres=TheApplication->GetSWApp()->get_IActiveDoc(&pModelDoc);//获得pModelDoc指针
	m_iSldWorks->get_IActiveDoc2(&pModelDoc);//获得pModelDoc指针

	pModelDoc->QueryInterface(IID_IPartDoc,(LPVOID *)&pPartDoc);//获得pPartDoc指针
	//part2.SLDPRT

	CString tempc=dlgx.m_Gczj;
	AfxMessageBox(tempc);//注意AfxMessageBox的_T格式

//因为是（米制）!

	double	m_L=atof(dlgx.m_L)/1000;
	double	m_B=atof(dlgx.m_B)/1000;
	double	m_H=atof(dlgx.m_H)/1000;
	double	m_h_=atof(dlgx.m_h_)/1000;
	double	m_h_1=atof(dlgx.m_h_1)/1000;
	double	m_h_2=atof(dlgx.m_h_2)/1000;
	double	m_C=atof(dlgx.m_C)/1000;
	double	m_l_=atof(dlgx.m_l_)/1000;
	double	m_l_1=atof(dlgx.m_l_1)/1000;
	double	m_l_2=atof(dlgx.m_l_2)/1000;
	double	m_l_3=atof(dlgx.m_l_3)/1000;
	double	m_b_=atof(dlgx.m_b_)/1000;
	double	m_b_1=atof(dlgx.m_b_1)/1000;
	double	m_r_=atof(dlgx.m_r_)/1000;
	double	m_r_1=atof(dlgx.m_r_1)/1000;

	CComBSTR c1_d1(_T("D1@草图1@bd05.SLDPRT"));//aha?
	CComBSTR c1_d2(_T("D2@草图1@bd05.SLDPRT"));//aha?
	CComBSTR c1_d3(_T("D3@草图1@bd05.SLDPRT"));//aha?
	CComBSTR c1_d4(_T("D4@草图1@bd05.SLDPRT"));//aha?
	CComBSTR c1_d5(_T("D5@草图1@bd05.SLDPRT"));//aha?
	CComBSTR c1_d6(_T("D6@草图1@bd05.SLDPRT"));//aha?

	CComBSTR tutai(_T("D1@凸台-拉伸1@bd05.SLDPRT"));//aha?

	CComBSTR c2_banjing(_T("b_/2@草图2@bd05.SLDPRT"));//aha?
	CComBSTR c2_l(_T("l_@草图2@bd05.SLDPRT"));//aha?
	CComBSTR c2_l1(_T("l_1@草图2@bd05.SLDPRT"));//aha?

	CComBSTR c7_2r(_T("2r_@草图7@bd05.SLDPRT"));//aha?

	CComBSTR c8_2h1(_T("2h_1@草图8@bd05.SLDPRT"));//aha?

	CComBSTR c10_l3(_T("l_3@草图10@bd05.SLDPRT"));////aha?

	CComBSTR c13_d2(_T("D2@草图13@bd05.SLDPRT"));////aha? //D2=(B-b1)/2



	pModelDoc->IParameter(c1_d1,&m_Dimension);	//为c1_d1(_T("D1@草图1@bd05.SLDPRT"))赋值
	m_Dimension->SetSystemValue2(m_L-m_l_3,swSetValue_UseCurrentSetting,&retval);

	pModelDoc->IParameter(c1_d2,&m_Dimension);
	m_Dimension->SetSystemValue2(m_h_2,swSetValue_UseCurrentSetting,&retval);

	pModelDoc->IParameter(c1_d3,&m_Dimension);
	m_Dimension->SetSystemValue2(m_L-m_l_2,swSetValue_UseCurrentSetting,&retval);

	pModelDoc->IParameter(c1_d4,&m_Dimension);
	m_Dimension->SetSystemValue2(m_H-m_h_2,swSetValue_UseCurrentSetting,&retval);

	pModelDoc->IParameter(c1_d5,&m_Dimension);
	m_Dimension->SetSystemValue2(m_l_2,swSetValue_UseCurrentSetting,&retval);

	pModelDoc->IParameter(c1_d6,&m_Dimension);
	m_Dimension->SetSystemValue2(m_H-m_h_,swSetValue_UseCurrentSetting,&retval);

	pModelDoc->IParameter(tutai,&m_Dimension);
	m_Dimension->SetSystemValue2(m_B,swSetValue_UseCurrentSetting,&retval);

	pModelDoc->IParameter(c2_banjing,&m_Dimension);
	m_Dimension->SetSystemValue2(m_b_/2,swSetValue_UseCurrentSetting,&retval);

	pModelDoc->IParameter(c2_l,&m_Dimension);
	m_Dimension->SetSystemValue2(m_l_,swSetValue_UseCurrentSetting,&retval);

	pModelDoc->IParameter(c2_l1,&m_Dimension);
	m_Dimension->SetSystemValue2(m_l_1,swSetValue_UseCurrentSetting,&retval);

	pModelDoc->IParameter(c7_2r,&m_Dimension);
	m_Dimension->SetSystemValue2(m_r_*2,swSetValue_UseCurrentSetting,&retval);

	pModelDoc->IParameter(c8_2h1,&m_Dimension);
	m_Dimension->SetSystemValue2(m_h_1*2,swSetValue_UseCurrentSetting,&retval);

	pModelDoc->IParameter(c10_l3,&m_Dimension);
	m_Dimension->SetSystemValue2(m_l_3,swSetValue_UseCurrentSetting,&retval);

	pModelDoc->IParameter(c13_d2,&m_Dimension);
	m_Dimension->SetSystemValue2(m_B/2-m_b_1/2,swSetValue_UseCurrentSetting,&retval);



	pModelDoc->BlankSketch ( );//清空所选的草图
	
	CComBSTR filename3(_T("*Isometric"));//aha?
	pModelDoc->ShowNamedView2(filename3,7);//显示等轴测视区
	pModelDoc->ViewZoomtofit2();//自动缩放以整屏显示全图
	pModelDoc->EditRebuild3(&retval_rebuild);

	//////////////////////////////////////////////////////////////////////////
	pModelDoc=NULL;//释放pModelDoc指针
	//pPartDoc->Release();//释放pPartDoc指针
	pPartDoc=NULL;
	}
	return S_OK;
}
