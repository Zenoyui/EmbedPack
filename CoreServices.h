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
    enum class ElementType : uint8_t
    {
        UnsignedChar = 0,
        Uint8,
        StdByte,
        UnsignedShort,
        Uint16,
        Uint32,
        Uint64
    };

    enum class ArrayStyle : uint8_t
    {
        ConstArray = 0,
        StaticConstArray,
        ConstexprArray,
        ConstexprStdArray,
        StaticConstexprStdArray
    };

    struct Format
    {
        ElementType elementType = ElementType::UnsignedChar;
        ArrayStyle arrayStyle   = ArrayStyle::ConstArray;
    };

    constexpr uint64_t UI_SOFT_LIMIT = 8ull * 1024ull * 1024ull;

    struct Job
    {
        HWND hwndNotify = nullptr;
        std::wstring inPath;
        std::wstring outPath;
        bool largeMode = false;
        Format format{};
    };

    bool GetFileSizeU64(const std::wstring& path, uint64_t& outSize);
    bool StartConversionAsync(const Job& job, std::wstring& outSmallOutput);
}
