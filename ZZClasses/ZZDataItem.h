#pragma once
#include <string>
class CZZDataItem
{
	std::string m_stringName;
public:
	std::string GetName() const { return m_stringName; }
	void SetName(std::string val) { m_stringName = val; }
	virtual std::string GetValueString() = 0;
	CZZDataItem(void);
	~CZZDataItem(void);
};
typedef CZZDataItem* PZZDataItem;

