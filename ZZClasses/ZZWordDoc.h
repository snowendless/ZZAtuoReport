#pragma once
#include <string>
#include <vector>
#include "ZZDataItem.h"
class CZZWordDoc
{
	std::wstring m_stringName;
	std::vector<PZZDataItem> m_vecDataItems;
public:
	HRESULT GenerateWordDoc(std::wstring templatePath,std::wstring LocationFolder);
	HRESULT AddDataItem(std::wstring DataName,std::wstring dataString);
	std::wstring GetName() const { return m_stringName; }
	void SetName(std::wstring val) { m_stringName = val; }
	CZZWordDoc(void);
	~CZZWordDoc(void);

	void ClearDataItem();

};
typedef CZZWordDoc* PZZWordDoc;
