#include "StdAfx.h"
#include "ZZExcel2Word.h"
#include <sstream>//for ostr<<_T("管道序号");
#include "..\CExcelApplication.h"
#include "..\CExcelWorkbook.h"
#include "..\CExcelWorkbooks.h"
#include "..\CExcelWorksheet.h"
#include "..\CExcelWorksheets.h"
#include "..\CExcelRange.h"
CZZExcel2Word::CZZExcel2Word(void)
{
	std::ostringstream   ostr;
	ostr<<_T("管道序号");
	m_stringWordDocKey =  ostr.str();
}


CZZExcel2Word::~CZZExcel2Word(void)
{
	ClearWordDoc();

}
HRESULT CZZExcel2Word::TransferExcelFiles2Word(std::vector<std::string> vecExcelFiles,std::string wordDocLocationFolder)
{
	return S_OK;
}
HRESULT CZZExcel2Word::BuildDataFromExcelFile(std::string ExcelFile,std::string stringDocKey)
{
	CExcelApplication ExcelApp; 
	CExcelWorkbooks wbsMyBooks;  
	CExcelWorkbook wbMyBook;  
	CExcelWorksheets wssMysheets; 
	CExcelWorksheet wsMysheet;  
	CExcelRange rgMyRge;  
	//创建Excel 2000服务器(启动Excel) 
	if (!ExcelApp.CreateDispatch(_T("Excel.Application"),NULL))  
	{   
		AfxMessageBox(_T("创建Excel服务失败!"));  
		exit(1);   
	}  

	//利用模板文件建立新文档  
	wbsMyBooks.AttachDispatch(ExcelApp.get_Workbooks(),true); 
	wbMyBook.AttachDispatch(wbsMyBooks.Add(_variant_t(_T("C:\\Users\\Administrator.UXEXD6YTDEVZ8JF\\Desktop\\太阳宫燃气热电管道（未做管线）台账整理2013042702new.xls")))); 
	//得到Worksheets   wssMysheets.AttachDispatch(wbMyBook.GetWorksheets(),true); 
	//得到sheet1   wsMysheet.AttachDispatch(wssMysheets.GetItem(_variant_t("sheet1")),true); 
	//释放对象  
	rgMyRge.ReleaseDispatch(); 
	wsMysheet.ReleaseDispatch();
	wssMysheets.ReleaseDispatch();
	wbMyBook.ReleaseDispatch();  
	wbsMyBooks.ReleaseDispatch();   
	ExcelApp.ReleaseDispatch();  

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
