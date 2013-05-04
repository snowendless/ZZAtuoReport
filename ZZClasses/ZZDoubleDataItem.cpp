#include "StdAfx.h"
#include "ZZDoubleDataItem.h"
#include <sstream>

CZZDoubleDataItem::CZZDoubleDataItem(void)
{
	m_dValue = 0;
}


CZZDoubleDataItem::~CZZDoubleDataItem(void)
{
}
 std::string CZZDoubleDataItem::GetValueString()
 {
	 std::ostringstream   ostr;

	 return ostr.str();
 }