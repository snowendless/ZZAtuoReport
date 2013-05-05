#include "StdAfx.h"
#include "ZZExcel2Word.h"
#include "..\CExcelApplication.h"
#include "..\CExcelWorkbook.h"
#include "..\CExcelWorkbooks.h"
#include "..\CExcelWorksheet.h"
#include "..\CExcelWorksheets.h"
#include "..\CExcelRange.h"
#include <sstream>
CZZExcel2Word::CZZExcel2Word(void)
{
	m_stringWordDocKey =  _T("管道序号");
	m_stringWordTemplatePath = _T("C:\\Users\\Administrator.UXEXD6YTDEVZ8JF\\AppData\\Roaming\\Microsoft\\Templates\\ZZTemplate.dot");
	m_iValueNameRow = 2;
}


CZZExcel2Word::~CZZExcel2Word(void)
{
	ClearWordDoc();

}
HRESULT CZZExcel2Word::ExportDataToWordDoc(std::wstring LocationFolder)
{
	std::vector<PZZWordDoc>::iterator it;

	for (it = m_vecWordDoc.begin(); it != m_vecWordDoc.end(); ++it)
	{
		PZZWordDoc temp = *it;
		temp->GenerateWordDoc(m_stringWordTemplatePath,LocationFolder);
		break;
	}
	return S_OK;
}

HRESULT CZZExcel2Word::TransferExcelFiles2Word(std::vector<std::wstring> vecExcelFiles,std::wstring wordDocLocationFolder)
{
	std::vector<std::wstring>::iterator it;

	for (it = vecExcelFiles.begin(); it != vecExcelFiles.end(); ++it)
	{
		std::wstring temp = *it;
		BuildDataFromExcelFile(temp,m_stringWordDocKey);
	}
	return S_OK;
}
static std::wstring GetStringFromExcelCell(CExcelRange& useRange)
{
	COleVariant keyValue = useRange.get_Value2();	
	std::wstring itemString;
	if (keyValue.vt != VT_BSTR)
	{
		if (keyValue.vt == VT_R8)
		{
			std::wostringstream ostr;
			ostr<<keyValue.dblVal;
			itemString = ostr.str();
		}
	}
	else
	{
		if (keyValue.bstrVal != NULL)
		{
			itemString = keyValue.bstrVal;
		}	
	}
	return itemString;
}
HRESULT CZZExcel2Word::BuildDataFromExcelFile(std::wstring ExcelFile,std::wstring stringDocKey)
{
	CExcelApplication ExcelApp; 
	CExcelWorkbooks wbsMyBooks;  
	CExcelWorkbook wbMyBook;  
	CExcelWorksheets wssMysheets; 
	CExcelWorksheet wsMysheet;  
	//创建Excel 2000服务器(启动Excel) 
	if (!ExcelApp.CreateDispatch(_T("Excel.Application"),NULL))  
	{   
		AfxMessageBox(_T("创建Excel服务失败!"));  
		exit(1);   
	}  

	//利用模板文件建立新文档  
	wbsMyBooks = ExcelApp.get_Workbooks(); 
	 COleVariant  avar((long)DISP_E_PARAMNOTFOUND,VT_ERROR); 
	wbMyBook = wbsMyBooks.Open(ExcelFile.c_str(),avar,avar,avar,avar,avar,avar,avar,avar,avar,avar,avar,avar,avar,avar); 
	//得到Worksheets   
	wssMysheets = wbMyBook.get_Sheets(); 
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND,   VT_ERROR);
	CExcelRange useRange;
	for (int iSheetIdx = 1; iSheetIdx <= wssMysheets.get_Count() ; iSheetIdx++)
	{
		//得到sheet 
		wsMysheet=wssMysheets.get_Item(COleVariant((short)iSheetIdx)); 
		
#ifdef _DEBUG
		CString sheetname = wsMysheet.get_Name();
#endif // _DEBUG
		//首先分析表的第一行，得到数据关键字
		int iKeyColumnIdx = -1;
		CExcelRange useRange = wsMysheet.get_UsedRange();

		long iRowNum = useRange.get_Count();
		useRange = useRange.get_Columns();

		long iStartRow = useRange.get_Row();
		int nColumn = useRange.get_Count();
		long iStartCol = useRange.get_Column();
		long iValueNameRow = m_iValueNameRow;
		//find keyName Column index
		for (int iColIdx = iStartCol; iColIdx <= nColumn ; iColIdx++)
		{
			useRange = wsMysheet.get_Cells();
			COleVariant keyValue=useRange.get_Item(_variant_t(iValueNameRow),_variant_t(iColIdx));
			useRange = keyValue.pdispVal;
			std::wstring itemString = GetStringFromExcelCell(useRange);
			if (stringDocKey == itemString )
			{
				iKeyColumnIdx = iColIdx;
				break;
			}
		}
		if (iKeyColumnIdx < 0)
		{
			continue;//no valid data with doc key string
		}
		//根据关键字列创建doc

		for (int iRodIdx = 3; iRodIdx <= iRowNum ; iRodIdx++)
		{
			useRange = wsMysheet.get_Cells();
			COleVariant keyValue=useRange.get_Item(_variant_t(iRodIdx),_variant_t(iKeyColumnIdx));
			useRange = keyValue.pdispVal;
			std::wstring keyitemString = GetStringFromExcelCell(useRange);
			if (keyitemString.empty())
			{
				continue;
			}
			PZZWordDoc pDoc = GetDocFromKeyString(keyitemString);
			if (pDoc == NULL)
			{
				pDoc = CreateDoc(keyitemString);
				if (pDoc == NULL)
				{
					continue;
				}
				m_vecWordDoc.push_back(pDoc);
			}

			//逐个读取相关doc的数据
			for (int iColIdx = iStartCol; iColIdx <= nColumn ; iColIdx++)
			{
				if (iColIdx == iKeyColumnIdx)
				{
					continue;
				}
				useRange = wsMysheet.get_Cells();
				keyValue=useRange.get_Item(_variant_t(iRodIdx),_variant_t(iColIdx));
				useRange = keyValue.pdispVal;
				std::wstring valueitemString  = GetStringFromExcelCell(useRange);
				if (valueitemString.empty())
				{
					//无效数据
					continue;
				}
				//查找这个值对应的名字
				useRange = wsMysheet.get_Cells();
				keyValue=useRange.get_Item(_variant_t(iValueNameRow),_variant_t(iColIdx));
				useRange = keyValue.pdispVal;	
				std::wstring valuenameitemString = GetStringFromExcelCell(useRange);

				if (valuenameitemString.empty())
				{
					//无效数据名字
					continue;
				}
				pDoc->AddDataItem(valuenameitemString,valueitemString);
			}
		}
		wsMysheet.ReleaseDispatch();  
	}//sheet scan
	wssMysheets.ReleaseDispatch();  
	wbsMyBooks.Close();
	wbMyBook.ReleaseDispatch();  
	wbsMyBooks.ReleaseDispatch();  

	ExcelApp.Quit();

	ExcelApp.ReleaseDispatch();
	return S_OK;
}


PZZWordDoc CZZExcel2Word::CreateDoc(std::wstring key)
{
	PZZWordDoc newDoc = new CZZWordDoc();
	newDoc->SetName(key);
	return newDoc;
}

PZZWordDoc CZZExcel2Word::GetDocFromKeyString(std::wstring key)
{
	std::vector<PZZWordDoc>::iterator it;

	for (it = m_vecWordDoc.begin(); it != m_vecWordDoc.end(); ++it)
	{
		PZZWordDoc temp = *it;
		if (temp->GetName() == key)
		{
			return temp;
		}
	}
	return NULL;
}

void CZZExcel2Word::ClearWordDoc()
{
	std::vector<PZZWordDoc>::iterator it;

	for (it = m_vecWordDoc.begin(); it != m_vecWordDoc.end(); ++it)
	{
		PZZWordDoc temp = *it;
		delete temp;
	}
	m_vecWordDoc.clear();
}
