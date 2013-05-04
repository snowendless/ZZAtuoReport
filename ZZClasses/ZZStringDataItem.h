#pragma once
#include "zzdataitem.h"
class CZZStringDataItem :
	public CZZDataItem
{
	std::string m_stringValue;
public:
	std::string GetStringValue() const { return m_stringValue; }
	void SetStringValue(std::string val) { m_stringValue = val; }
	virtual std::string GetValueString(){ return m_stringValue; }
	CZZStringDataItem(void);
	~CZZStringDataItem(void);
};

