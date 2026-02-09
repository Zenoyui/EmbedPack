// main.cpp
#include "App.h"

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, PWSTR, int nCmdShow)
{
    EmbedPack::App app(hInstance);
    return app.Run(nCmdShow);
}
