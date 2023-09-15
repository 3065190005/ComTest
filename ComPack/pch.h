#ifndef PCH_H
#define PCH_H

#include "framework.h"
#include <combaseapi.h>
extern "C" _declspec(dllimport) 
VARIANT CallByString(
	const WCHAR * __comname, 
	const WCHAR * __funcname, 
	int __count, 
	...);

extern "C" _declspec(dllimport)
VARIANT CallByDispatch(
	IDispatch* object,
	const WCHAR * __funcname,
	int __count,
	...);

extern "C" _declspec(dllimport)
VARIANT CallByUnknown(
	IUnknown * object,
	const WCHAR * __funcname,
	int __count,
	...);

#endif //PCH_H
