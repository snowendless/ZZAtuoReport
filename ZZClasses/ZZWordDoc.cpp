#include "StdAfx.h"
#include "ZZWordDoc.h"


CZZWordDoc::CZZWordDoc(void)
{
}


CZZWordDoc::~CZZWordDoc(void)
{
	ClearDataItem();

	
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
