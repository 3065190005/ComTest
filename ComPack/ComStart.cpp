#include "pch.h"
#include "ComStart.h"

bool Init()
{
    if (_init == true)
    {
        return _init;
    }
    else
    {
        if (S_OK == CoInitialize(NULL))
            _init = true;
        else
            _init = false;

        return _init;
    }

    return false;
}

void Release()
{
    if (true == _init)
    {
        CoUninitialize();
        _init = false;
    }
}
