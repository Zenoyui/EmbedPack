// CoreServices.h
#pragma once

#define NOMINMAX
#include <windows.h>

#include <cstdint>
#include <string>

namespace EmbedPack::AppMessages
{
    constexpr UINT WM_APP_PROGRESS = WM_APP + 1;
    constexpr UINT WM_APP_DONE     = WM_APP + 2;
}

namespace EmbedPack::Clipboard
{
    void SetClipboardUnicode(HWND owner, const std::wstring& text);
}

namespace EmbedPack::FileDialogs
{
    bool PromptOpenInputFile(HWND owner, std::wstring& outPath);
    bool PromptSaveOutputPath(HWND owner, const std::wstring& inputPath, std::wstring& outPath);
}

namespace EmbedPack::Converter
{
    constexpr uint64_t UI_SOFT_LIMIT = 8ull * 1024ull * 1024ull;

    struct Job
    {
        HWND hwndNotify = nullptr;
        std::wstring inPath;
        std::wstring outPath;
        bool largeMode = false;
    };

    bool GetFileSizeU64(const std::wstring& path, uint64_t& outSize);
    bool StartConversionAsync(const Job& job, std::wstring& outSmallOutput);
}
