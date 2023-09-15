#include "pch.h"
#include "ComStart.h"
#include <assert.h>
#include <atlbase.h>

// 报错机制
#define ASSERT(s) if((s) == true)

// Com类名(MyCom.FirstClass)，函数名(Add)，传入的参数数量(2)，变长参数[VARIANT](VARIANT.vt=VT_I4 VARIANT.lVar = 2)
VARIANT CallByString(const WCHAR* __comname,const WCHAR* __funcname,int __count, ...)
{
    /* Com注册到系统后使用 */
   
    // 是否成功初始化
    if (true == Init())
    {
        // ProgId值存放
        CLSID clsid;

        // 通过 ProgID 取得组件的 CLSID
        // CLSID 值存放在注册表 HKEY_CLASSES_ROOT [以__comname加.1为键值(MyCom.FirstClass.1)]
        HRESULT hr = ::CLSIDFromProgID(__comname, &clsid);

        ASSERT(S_OK != hr)
            assert(hr != S_OK);

        // 智能指针获取 IUnknow
        CComPtr<IUnknown>spUnk;

        /* 
        *   CoCreateInstance
        *       CLSIDFromProgId获取的值
        *       指向接口IUnknown的指针
        *       运行可执行代码的上下文[CLSCTX_ALL 为所有]
        *       IID_IUnknown为返回类型
        *       用来接收指向Com对象接口地址的指针变量
        */
        // 获取IUnknow内容 
        hr = ::CoCreateInstance(clsid, NULL, CLSCTX_ALL, IID_IUnknown, (LPVOID*)&spUnk);

        ASSERT(S_OK != hr)
            assert(hr != S_OK);
        
        // 通过IUnknown类型智能指针，获取IDispatch类型智能指针
        CComDispatchDriver spDisp(spUnk);

        // 参数数组
        VARIANT* __args = new VARIANT[__count];

        // 变长参数变量
        va_list ap;

        // 定位到第一个函数变长参数
        va_start(ap, __count);

        // 循环获取变长参数，并转换为 VARIANT 类型放入 __args变量
        for (auto i = 0; i < __count; i++)
            __args[__count -1 - i] = va_arg(ap, VARIANT);

        // 结束变长参数
        va_end(ap);

        // Com函数返回值存放
        VARIANT __ret;

        // 执行Com函数

        /*
        *   [InvokeN]
        *       函数名
        *       函数参数
        *       函数数量
        *       返回值存放处
        */
        hr = spDisp.InvokeN((LPCOLESTR)__funcname, __args, __count, &__ret);
       
        ASSERT(S_OK != hr)
            assert(hr != S_OK);

        // 内存回收
        delete[] __args;

        // 返回值
        return __ret;
    }


    // 卸载 Com库
    Release();

    assert(_init == false);
}

VARIANT CallByDispatch(IDispatch* object, const WCHAR* __funcname, int __count, ...)
{
    /* Com注册到系统后使用 */

// 是否成功初始化
    if (true == Init())
    {
        HRESULT hr = NULL;
        if (object != nullptr)
            hr = S_OK;
        // 通过IDispatch指针，获取IDispatch类型智能指针
        CComDispatchDriver spDisp(object);

        // 参数数组
        VARIANT* __args = new VARIANT[__count];

        // 变长参数变量
        va_list ap;

        // 定位到第一个函数变长参数
        va_start(ap, __count);

        // 循环获取变长参数，并转换为 VARIANT 类型放入 __args变量
        for (auto i = 0; i < __count; i++)
            __args[__count - 1 - i] = va_arg(ap, VARIANT);

        // 结束变长参数
        va_end(ap);

        // Com函数返回值存放
        VARIANT __ret;

        // 执行Com函数

        /*
        *   [InvokeN]
        *       函数名
        *       函数参数
        *       函数数量
        *       返回值存放处
        */
        hr = spDisp.InvokeN((LPCOLESTR)__funcname, __args, __count, &__ret);

        ASSERT(S_OK != hr)
            assert(hr != S_OK);

        // 内存回收
        delete[] __args;

        // 返回值
        return __ret;
    }


    // 卸载 Com库
    Release();

    assert(_init == false);
}

VARIANT CallByUnknown(IUnknown* object, const WCHAR* __funcname, int __count, ...)
{
    /* Com注册到系统后使用 */

// 是否成功初始化
    if (true == Init())
    {
        HRESULT hr = NULL;
        if (object != nullptr)
            hr = S_OK;

        // 通过IUnknown指针，获取IDispatch类型智能指针
        CComDispatchDriver spDisp(object);

        // 参数数组
        VARIANT* __args = new VARIANT[__count];

        // 变长参数变量
        va_list ap;

        // 定位到第一个函数变长参数
        va_start(ap, __count);

        // 循环获取变长参数，并转换为 VARIANT 类型放入 __args变量
        for (auto i = 0; i < __count; i++)
            __args[__count - 1 - i] = va_arg(ap, VARIANT);

        // 结束变长参数
        va_end(ap);

        // Com函数返回值存放
        VARIANT __ret;

        // 执行Com函数

        /*
        *   [InvokeN]
        *       函数名
        *       函数参数
        *       函数数量
        *       返回值存放处
        */
        hr = spDisp.InvokeN((LPCOLESTR)__funcname, __args, __count, &__ret);

        ASSERT(S_OK != hr)
            assert(hr != S_OK);

        // 内存回收
        delete[] __args;

        // 返回值
        return __ret;
    }


    // 卸载 Com库
    Release();

    assert(_init == false);
}