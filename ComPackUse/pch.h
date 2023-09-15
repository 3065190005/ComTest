// pch.h

#pragma once
#include <combaseapi.h>

// 新添加的代码
#pragma comment(lib,"ComPack.lib")

extern "C" _declspec(dllimport) VARIANT CallByString(const WCHAR * __comname, const WCHAR * __funcname, int __count, ...);
extern "C" _declspec(dllimport) VARIANT CallByDispatch(IDispatch * object, const WCHAR * __funcname, int __count, ...);
extern "C" _declspec(dllimport) VARIANT CallByUnknown(IUnknown * object, const WCHAR * __funcname, int __count, ...);
