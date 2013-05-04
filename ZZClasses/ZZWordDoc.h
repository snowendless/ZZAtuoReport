#pragma once
#include <string>
#include <vector>
#include "ZZDataItem.h"
class CZZWordDoc
{
	std::string m_stringName;
	std::vector<PZZDataItem> m_vecDataItems;
public:
	std::string GetName() const { return m_stringName; }
	void SetName(std::string val) { m_stringName = val; }
	CZZWordDoc(void);
	~CZZWordDoc(void);

	void ClearDataItem();

};
typedef CZZWordDoc* PZZWordDoc;
