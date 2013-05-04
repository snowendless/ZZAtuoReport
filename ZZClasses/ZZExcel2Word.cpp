#include "StdAfx.h"
#include "ZZExcel2Word.h"
#include <sstream>//for ostr<<_T("�ܵ����");
#include "..\CExcelApplication.h"
#include "..\CExcelWorkbook.h"
#include "..\CExcelWorkbooks.h"
#include "..\CExcelWorksheet.h"
#include "..\CExcelWorksheets.h"
#include "..\CExcelRange.h"
CZZExcel2Word::CZZExcel2Word(void)
{
	std::ostringstream   ostr;
	ostr<<_T("�ܵ����");
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
	//����Excel 2000������(����Excel) 
	if (!ExcelApp.CreateDispatch(_T("Excel.Application"),NULL))  
	{   
		AfxMessageBox(_T("����Excel����ʧ��!"));  
		exit(1);   
	}  

	//����ģ���ļ��������ĵ�  
	wbsMyBooks.AttachDispatch(ExcelApp.get_Workbooks(),true); 
	wbMyBook.AttachDispatch(wbsMyBooks.Add(_variant_t(_T("C:\\Users\\Administrator.UXEXD6YTDEVZ8JF\\Desktop\\̫����ȼ���ȵ�ܵ���δ�����ߣ�̨������2013042702new.xls")))); 
	//�õ�Worksheets   wssMysheets.AttachDispatch(wbMyBook.GetWorksheets(),true); 
	//�õ�sheet1   wsMysheet.AttachDispatch(wssMysheets.GetItem(_variant_t("sheet1")),true); 
	//�ͷŶ���  
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
