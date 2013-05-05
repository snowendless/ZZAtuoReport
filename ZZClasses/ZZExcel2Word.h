#pragma once
#include <string>
#include <vector>
#include "ZZWordDoc.h"
class CZZExcel2Word
{
	std::wstring m_stringWordDocKey;
	std::wstring m_stringWordTemplatePath;
	std::vector<PZZWordDoc> m_vecWordDoc;
	PZZWordDoc GetDocFromKeyString(std::wstring key);
	PZZWordDoc CreateDoc(std::wstring key);
	void ClearWordDoc();
public:
	std::wstring GetStringWordTemplatePath() const { return m_stringWordTemplatePath; }
	void SetStringWordTemplatePath(std::wstring val) { m_stringWordTemplatePath = val; }
	HRESULT ExportDataToWordDoc(std::wstring LocationFolder);
	HRESULT TransferExcelFiles2Word(std::vector<std::wstring> vecExcelFiles,std::wstring wordDocLocationFolder);
	HRESULT BuildDataFromExcelFile(std::wstring ExcelFile,std::wstring stringDocKey);
	CZZExcel2Word(void);
	~CZZExcel2Word(void);


};

