#include <iostream>
#include "pch.h"


int main()
{
	VARIANT __param1,__param2,__param3;

	VariantInit(&__param1);
	VariantInit(&__param2);
	VariantInit(&__param3);

	// 参数类型为 string/wstring 值为 d:\\testfile.txt
	__param1.vt = VT_BSTR;
	__param1.bstrVal = SysAllocString(L"d:\\testfile.txt");

	// 参数类型为 bool 值为 true
	__param2.vt = VT_BOOL;
	__param2.boolVal = VARIANT_TRUE;

	// 参数类型为 bool 值为 true
	__param3.vt = VT_BOOL;
	__param3.boolVal = VARIANT_TRUE;

	VARIANT __ret1 = CallByString(L"Scripting.FileSystemObject", L"CreateTextFile",2, __param1, __param2);
	IUnknown* ptr = (IUnknown*)__ret1.pdispVal;
	
	VARIANT __ret2 = CallByDispatch(__ret1.pdispVal, L"WriteLine",1, __param1);
	VARIANT __ret3 = CallByDispatch(__ret1.pdispVal, L"WriteLine",1, __param1);
	VARIANT __ret4 = CallByUnknown(ptr, L"Close",0);




	VariantClear(&__param1);
	VariantClear(&__param2);
	VariantClear(&__param3);

	VariantClear(&__ret1);
	VariantClear(&__ret2);
	VariantClear(&__ret3);
	VariantClear(&__ret4);

}


/*
* 
* let object = null;
* object = ComCall("Scripting.FileSystemObject","CreateTextFile",["d:\\testfile.txt",true]);
* ComCall(object,"WriteLine",["line 1"]);
* ComCall(object,"WriteLine",["line 2"]);
* 
* ComFree(object);
*/
