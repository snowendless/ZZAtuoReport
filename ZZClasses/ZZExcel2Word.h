#pragma once
#include <string>
#include <vector>
#include "ZZWordDoc.h"
class CZZExcel2Word
{
	std::string m_stringWordDocKey;
	std::vector<PZZWordDoc> m_vecWordDoc;

public:
	HRESULT TransferExcelFiles2Word(std::vector<std::string> vecExcelFiles,std::string wordDocLocationFolder);
	HRESULT BuildDataFromExcelFile(std::string ExcelFile,std::string stringDocKey);
	CZZExcel2Word(void);
	~CZZExcel2Word(void);

	void ClearWordDoc();

};

