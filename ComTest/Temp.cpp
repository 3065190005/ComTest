// Temp.cpp: CTemp 的实现

#include "pch.h"
#include "Temp.h"


// CTemp

STDMETHODIMP CTemp::Number(LONG __num, LONG* __result)
{
	
	// STDMETHODIMP 宏 HRESULT __stdcall 
	// __result 是返回值
	// __num 为参数
	*__result = __num * __num;
	
	// 默认返回 函数正确执行
	return S_OK;
}
