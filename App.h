// App.h
#pragma once

#define NOMINMAX
#include <windows.h>

namespace EmbedPack
{
    class App final
    {
    public:
        explicit App(HINSTANCE hInstance);
        int Run(int nCmdShow);

    private:
        HINSTANCE m_hInstance = nullptr;
    };
}
