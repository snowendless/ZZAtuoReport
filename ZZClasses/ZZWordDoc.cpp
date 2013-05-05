#include "StdAfx.h"
#include "ZZWordDoc.h"
#include "ZZStringDataItem.h"
#include "..\CWordApplication.h"
#include "..\CWordDocuments.h"
#include "..\CWordDocument.h"
#include "..\CWordBookmarks.h"
#include "..\CWordBookmark.h"
#include "..\CWordRange.h"
CZZWordDoc::CZZWordDoc(void)
{
}


CZZWordDoc::~CZZWordDoc(void)
{
	ClearDataItem();

	
}
HRESULT CZZWordDoc::GenerateWordDoc(std::wstring templatePath,std::wstring LocationFolder)
{
	CWordApplication WordApp; 

	if (!WordApp.CreateDispatch(_T("Word.Application"),NULL))  
	{   
		AfxMessageBox(_T("创建Word服务失败!"));  
		exit(1);   
	}  
	CWordDocuments WordDocuments = WordApp.get_Documents();
	CComVariant tpl(templatePath.c_str()),Visble,DocType(0),NewTemplate(false);
	CWordDocument wordProduct=WordDocuments.Add(&tpl,&NewTemplate,&DocType,&Visble);

	
	CComVariant FileName((LocationFolder + m_stringName).c_str()); //文件名
	//TODO:保存WORD
	CComVariant FileFormat(0),LockComments(false),Password(_T("")),AddToRecentFiles(true),WritePassword(_T(""));
	CComVariant ReadOnlyRecommended(false),EmbedTrueTypeFonts(false),SaveNativePictureFormat(false),SaveFormsData(false),SaveAsAOCELetter(false);
	CComVariant Encoding(false),InsertLineBreaks(false),AllowSubstitutions(false),LineEnding(false),AddBiDiMarks(false);

	//插入书签数据
	CWordBookmarks t_myBookMarks = wordProduct.get_Bookmarks();
	std::vector<PZZDataItem>::iterator it;

	for (it = m_vecDataItems.begin(); it != m_vecDataItems.end(); ++it)
	{
		PZZDataItem temp = *it;
		if (temp == NULL || temp->GetName().empty() || temp->GetValueString().empty())
		{
			continue;
		}
		try
		{
			CWordBookmark t_bookMark = t_myBookMarks.Item(COleVariant(temp->GetName().c_str()));
			CWordRange tBMRange = t_bookMark.get_Range();
			tBMRange.put_Text(temp->GetValueString().c_str());
			tBMRange.ReleaseDispatch();
			t_bookMark.ReleaseDispatch();
		}
		catch (...)
		{
			continue;
		}
	}

	wordProduct.SaveAs(&FileName,&FileFormat,&LockComments,&Password,&AddToRecentFiles,&WritePassword,
		&ReadOnlyRecommended,&EmbedTrueTypeFonts,&SaveNativePictureFormat,&SaveFormsData, &SaveAsAOCELetter,
		&Encoding,&InsertLineBreaks,&AllowSubstitutions,&LineEnding,&AddBiDiMarks);
	CComVariant saveChanges(false),OriginalFormat,RouteDocument;
	WordApp.Quit(&saveChanges,&OriginalFormat,&RouteDocument);
	t_myBookMarks.ReleaseDispatch();
	wordProduct.ReleaseDispatch();
	WordDocuments.ReleaseDispatch();
	WordApp.ReleaseDispatch();
	return S_OK;
}

HRESULT CZZWordDoc::AddDataItem(std::wstring DataName,std::wstring dataString)
{
	PZZStringDataItem pStringData = new CZZStringDataItem;
	PZZDataItem pData = dynamic_cast<PZZDataItem>(pStringData);
	if (pData == NULL)
	{
		delete pStringData;
		return E_FAIL;
	}
	pStringData->SetName(DataName);
	pStringData->SetStringValue(dataString);
	m_vecDataItems.push_back(pData);
	return S_OK;
}
void CZZWordDoc::ClearDataItem()
{
	std::vector<PZZDataItem>::iterator it;

	for (it = m_vecDataItems.begin(); it != m_vecDataItems.end(); ++it)
	{
		PZZDataItem temp = *it;
		delete temp;
	}
	m_vecDataItems.clear();
}
