// CoreServices.cpp
#include "CoreServices.h"

#include <commdlg.h>

#include <cstring>
#include <cwchar>
#include <algorithm>
#include <limits>
#include <string>
#include <utility>
#include <vector>

namespace
{
    std::wstring GetBaseNameNoExt(const std::wstring& path)
    {
        const size_t slash = path.find_last_of(L"\\/");
        const size_t start = (slash == std::wstring::npos) ? 0u : (slash + 1u);

        size_t dot = path.find_last_of(L'.');
        if (dot == std::wstring::npos || dot < start)
            dot = path.size();

        return path.substr(start, dot - start);
    }

    std::wstring GetDirName(const std::wstring& path)
    {
        const size_t slash = path.find_last_of(L"\\/");
        if (slash == std::wstring::npos)
            return L"";
        return path.substr(0, slash + 1u);
    }
}

namespace EmbedPack::Clipboard
{
    void SetClipboardUnicode(HWND owner, const std::wstring& text)
    {
        if (!OpenClipboard(owner))
            return;

        EmptyClipboard();

        const size_t bytes = (text.size() + 1u) * sizeof(wchar_t);
        HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, bytes);
        if (hMem != nullptr)
        {
            void* pMem = GlobalLock(hMem);
            if (pMem != nullptr)
            {
                std::memcpy(pMem, text.c_str(), bytes);
                GlobalUnlock(hMem);

                SetClipboardData(CF_UNICODETEXT, hMem);
                hMem = nullptr; // ownership передано системе
            }

            if (hMem != nullptr)
                GlobalFree(hMem);
        }

        CloseClipboard();
    }
}

namespace EmbedPack::FileDialogs
{
    bool PromptOpenInputFile(HWND owner, std::wstring& outPath)
    {
        outPath.clear();

        OPENFILENAMEW ofn{};
        wchar_t szFile[MAX_PATH]{};

        ofn.lStructSize = sizeof(ofn);
        ofn.hwndOwner = owner;
        ofn.lpstrFile = szFile;
        ofn.lpstrFile[0] = L'\0';
        ofn.nMaxFile = static_cast<DWORD>(std::size(szFile));
        ofn.lpstrFilter = L"All Files (*.*)\0*.*\0";
        ofn.nFilterIndex = 1;
        ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_EXPLORER;

        if (!GetOpenFileNameW(&ofn))
            return false;

        outPath = ofn.lpstrFile;
        return !outPath.empty();
    }

    bool PromptSaveOutputPath(HWND owner, const std::wstring& inputPath, std::wstring& outPath)
    {
        outPath.clear();

        const std::wstring dir = GetDirName(inputPath);
        const std::wstring base = GetBaseNameNoExt(inputPath);
        const std::wstring suggested = dir + base + L"_bytes.h";

        wchar_t buf[MAX_PATH]{};
        wcsncpy_s(buf, suggested.c_str(), _TRUNCATE);

        OPENFILENAMEW ofn{};
        ofn.lStructSize = sizeof(ofn);
        ofn.hwndOwner = owner;
        ofn.lpstrFile = buf;
        ofn.nMaxFile = static_cast<DWORD>(std::size(buf));
        ofn.lpstrFilter =
            L"C/C++ Header (*.h)\0*.h\0"
            L"C/C++ Source (*.cpp)\0*.cpp\0"
            L"All Files (*.*)\0*.*\0";
        ofn.nFilterIndex = 1;
        ofn.Flags = OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT | OFN_EXPLORER;
        ofn.lpstrDefExt = L"h";

        if (!GetSaveFileNameW(&ofn))
            return false;

        outPath = buf;
        return !outPath.empty();
    }
}

namespace EmbedPack::Converter
{
    namespace
    {
        struct Handle final
        {
            HANDLE h = nullptr;

            Handle() = default;
            explicit Handle(HANDLE hh) : h(hh) {}

            ~Handle()
            {
                if (h != nullptr && h != INVALID_HANDLE_VALUE)
                    CloseHandle(h);
            }

            Handle(const Handle&) = delete;
            Handle& operator=(const Handle&) = delete;

            Handle(Handle&& o) noexcept : h(o.h) { o.h = nullptr; }
            Handle& operator=(Handle&& o) noexcept
            {
                if (this != &o)
                {
                    if (h != nullptr && h != INVALID_HANDLE_VALUE)
                        CloseHandle(h);
                    h = o.h;
                    o.h = nullptr;
                }
                return *this;
            }

            bool valid() const noexcept { return (h != nullptr && h != INVALID_HANDLE_VALUE); }
            operator HANDLE() const noexcept { return h; }
        };

        struct MappedView final
        {
            const void* p = nullptr;

            MappedView() = default;
            explicit MappedView(const void* pp) : p(pp) {}

            ~MappedView()
            {
                if (p != nullptr)
                    UnmapViewOfFile(p);
            }

            MappedView(const MappedView&) = delete;
            MappedView& operator=(const MappedView&) = delete;

            MappedView(MappedView&& o) noexcept : p(o.p) { o.p = nullptr; }
            MappedView& operator=(MappedView&& o) noexcept
            {
                if (this != &o)
                {
                    if (p != nullptr)
                        UnmapViewOfFile(p);
                    p = o.p;
                    o.p = nullptr;
                }
                return *this;
            }

            const void* get() const noexcept { return p; }
            bool valid() const noexcept { return p != nullptr; }
        };

        static constexpr char HEXA[16] = {
            '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F'
        };

        struct FormatSpec
        {
            Converter::ElementType type = Converter::ElementType::UnsignedChar;
            const char* typeName = "unsigned char";
            size_t elemSize = 1u;
            bool needsCstdint = false;
            bool needsCstddef = false;
            bool usesStdByte = false;
        };

        static FormatSpec GetFormatSpec(Converter::ElementType t)
        {
            using Converter::ElementType;
            switch (t)
            {
            case ElementType::UnsignedChar:  return { t, "unsigned char", 1u, false, true,  false };
            case ElementType::Uint8:         return { t, "uint8_t",       1u, true,  false, false };
            case ElementType::StdByte:       return { t, "std::byte",     1u, false, true,  true  };
            case ElementType::UnsignedShort: return { t, "unsigned short",2u, false, true,  false };
            case ElementType::Uint16:        return { t, "uint16_t",      2u, true,  false, false };
            case ElementType::Uint32:        return { t, "uint32_t",      4u, true,  false, false };
            case ElementType::Uint64:        return { t, "uint64_t",      8u, true,  false, false };
            default:                         return { ElementType::UnsignedChar, "unsigned char", 1u, false, true,  false };
            }
        }

        struct StyleSpec
        {
            Converter::ArrayStyle style = Converter::ArrayStyle::ConstArray;
            const char* prefixNonArray = "const ";
            const char* prefixStdArray = "const ";
            const char* sizeQualifier  = "const ";
            bool usesStdArray = false;
        };

        static StyleSpec GetStyleSpec(Converter::ArrayStyle s)
        {
            using Converter::ArrayStyle;
            switch (s)
            {
            case ArrayStyle::ConstArray:
                return { s, "const ", "const ", "const ", false };
            case ArrayStyle::StaticConstArray:
                return { s, "static const ", "static const ", "static const ", false };
            case ArrayStyle::ConstexprArray:
                return { s, "constexpr ", "constexpr ", "constexpr ", false };
            case ArrayStyle::ConstexprStdArray:
                return { s, "", "constexpr ", "constexpr ", true };
            case ArrayStyle::StaticConstexprStdArray:
                return { s, "", "static constexpr ", "static constexpr ", true };
            default:
                return { ArrayStyle::ConstArray, "const ", "const ", "const ", false };
            }
        }

        static size_t ValuesPerLine(size_t elemSize)
        {
            const size_t v = (elemSize == 0u) ? 1u : (16u / elemSize);
            return (v == 0u) ? 1u : v;
        }

        static bool WriteAll(HANDLE h, const void* data, DWORD size, std::wstring& err)
        {
            const uint8_t* p = static_cast<const uint8_t*>(data);
            DWORD remaining = size;
            while (remaining > 0u)
            {
                DWORD written = 0u;
                if (!WriteFile(h, p, remaining, &written, nullptr))
                {
                    err = L"Failed to write output file.";
                    return false;
                }
                if (written == 0u)
                {
                    err = L"Failed to write output file (0 bytes written).";
                    return false;
                }
                p += written;
                remaining -= written;
            }
            return true;
        }

        static void AppendHexValue(std::string& dst, uint64_t value, size_t hexDigits)
        {
            dst.push_back('0');
            dst.push_back('x');
            for (size_t i = 0u; i < hexDigits; ++i)
            {
                const size_t shift = (hexDigits - 1u - i) * 4u;
                const uint8_t nibble = static_cast<uint8_t>((value >> shift) & 0x0Fu);
                dst.push_back(HEXA[nibble]);
            }
        }

        static void AppendIncludes(const FormatSpec& f, const StyleSpec& s, std::string& out)
        {
            bool any = false;
            if (f.needsCstdint)
            {
                out.append("#include <cstdint>\r\n");
                any = true;
            }
            if (f.needsCstddef || s.usesStdArray || f.needsCstdint)
            {
                out.append("#include <cstddef>\r\n");
                any = true;
            }
            if (s.usesStdArray)
            {
                out.append("#include <array>\r\n");
                any = true;
            }

            if (any)
                out.append("\r\n");
        }

        static void AppendHeader(const FormatSpec& f, const StyleSpec& s, size_t elementCount, std::string& out)
        {
            if (s.usesStdArray)
            {
                out.append(s.prefixStdArray);
                out.append("std::array<");
                out.append(f.typeName);
                out.append(", ");
                out.append(std::to_string(elementCount));
                out.append("> fileBytes = {");
            }
            else
            {
                out.append(s.prefixNonArray);
                out.append(f.typeName);
                out.append(" fileBytes[] = {");
            }
        }

        static void AppendFooter(
            const FormatSpec& f,
            const StyleSpec& s,
            size_t elementCount,
            size_t byteCount,
            std::string& out)
        {
            out.append("\r\n};\r\n");

            out.append(s.sizeQualifier);
            out.append("size_t fileBytesSize = ");
            out.append("sizeof(fileBytes);\r\n");

            const size_t paddedBytes = elementCount * f.elemSize;
            if (paddedBytes != byteCount)
            {
                out.append(s.sizeQualifier);
                out.append("size_t fileBytesOriginalSize = ");
                out.append(std::to_string(byteCount));
                out.append(";\r\n");
            }
        }

        static void AppendValueToken(const FormatSpec& f, uint64_t value, size_t hexDigits, std::string& out)
        {
            if (f.usesStdByte)
                out.append("std::byte{");

            AppendHexValue(out, value, hexDigits);

            if (f.usesStdByte)
                out.push_back('}');
        }

        static void BuildArrayAscii(
            const uint8_t* data,
            size_t byteCount,
            const Converter::Format& fmt,
            std::string& out)
        {
            const FormatSpec f = GetFormatSpec(fmt.elementType);
            const StyleSpec s = GetStyleSpec(fmt.arrayStyle);

            const size_t elemSize = f.elemSize;
            const size_t elementCount = (elemSize == 0u)
                ? 0u
                : ((byteCount + elemSize - 1u) / elemSize);

            const size_t hexDigits = std::max<size_t>(2u, elemSize * 2u);
            const size_t valuesPerLine = ValuesPerLine(elemSize);

            out.clear();
            out.reserve(byteCount * 5u + 256u);

            AppendIncludes(f, s, out);
            AppendHeader(f, s, elementCount, out);

            for (size_t i = 0u; i < elementCount; ++i)
            {
                if ((i % valuesPerLine) == 0u)
                    out.append("\r\n    ");

                uint64_t value = 0u;
                const size_t base = i * elemSize;
                for (size_t b = 0u; b < elemSize; ++b)
                {
                    const size_t idx = base + b;
                    const uint8_t byte = (idx < byteCount) ? data[idx] : 0u;
                    value |= static_cast<uint64_t>(byte) << (8u * b);
                }

                AppendValueToken(f, value, hexDigits, out);

                if (i + 1u != elementCount)
                    out.append(", ");
            }

            AppendFooter(f, s, elementCount, byteCount, out);
        }

        static bool ConvertSmallToMemory(
            const std::wstring& path,
            const Converter::Format& fmt,
            std::wstring& out,
            std::wstring& err)
        {
            err.clear();
            out.clear();

            Handle hFile(CreateFileW(
                path.c_str(),
                GENERIC_READ,
                FILE_SHARE_READ,
                nullptr,
                OPEN_EXISTING,
                FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN,
                nullptr));

            if (!hFile.valid())
            {
                err = L"Failed to open the file.";
                return false;
            }

            LARGE_INTEGER liSize{};
            if (!GetFileSizeEx(hFile, &liSize))
            {
                err = L"Failed to query file size.";
                return false;
            }

            if (liSize.QuadPart < 0)
            {
                err = L"Invalid file size.";
                return false;
            }

            const uint64_t fileSize64 = static_cast<uint64_t>(liSize.QuadPart);
            if (fileSize64 > static_cast<uint64_t>(std::numeric_limits<size_t>::max()))
            {
                err = L"File is too large for this process.";
                return false;
            }

            const size_t fileSize = static_cast<size_t>(fileSize64);

            Handle hMap(CreateFileMappingW(hFile, nullptr, PAGE_READONLY, 0u, 0u, nullptr));
            if (!hMap.valid())
            {
                err = L"Failed to create file mapping.";
                return false;
            }

            MappedView view(MapViewOfFile(hMap, FILE_MAP_READ, 0u, 0u, 0u));
            if (!view.valid() && fileSize != 0u)
            {
                err = L"Failed to map file view.";
                return false;
            }

            const auto* data = static_cast<const uint8_t*>(view.get());

            std::string ascii;
            BuildArrayAscii(data, fileSize, fmt, ascii);
            out.assign(ascii.begin(), ascii.end());
            return true;
        }

        static bool ConvertLargeToFile(
            const std::wstring& inPath,
            const std::wstring& outPath,
            HWND notifyHwnd,
            const Converter::Format& fmt,
            std::wstring& err)
        {
            err.clear();

            Handle hIn(CreateFileW(
                inPath.c_str(),
                GENERIC_READ,
                FILE_SHARE_READ,
                nullptr,
                OPEN_EXISTING,
                FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN,
                nullptr));

            if (!hIn.valid())
            {
                err = L"Failed to open the input file.";
                return false;
            }

            LARGE_INTEGER liSize{};
            if (!GetFileSizeEx(hIn, &liSize))
            {
                err = L"Failed to query input file size.";
                return false;
            }

            if (liSize.QuadPart < 0)
            {
                err = L"Invalid input file size.";
                return false;
            }

            const uint64_t fileSize64 = static_cast<uint64_t>(liSize.QuadPart);
            if (fileSize64 > static_cast<uint64_t>(std::numeric_limits<size_t>::max()))
            {
                err = L"File is too large for this process.";
                return false;
            }

            const size_t fileSize = static_cast<size_t>(fileSize64);

            Handle hMap(CreateFileMappingW(hIn, nullptr, PAGE_READONLY, 0u, 0u, nullptr));
            if (!hMap.valid())
            {
                err = L"Failed to create file mapping.";
                return false;
            }

            MappedView view(MapViewOfFile(hMap, FILE_MAP_READ, 0u, 0u, 0u));
            if (!view.valid() && fileSize != 0u)
            {
                err = L"Failed to map file view.";
                return false;
            }

            Handle hOut(CreateFileW(
                outPath.c_str(),
                GENERIC_WRITE,
                0,
                nullptr,
                CREATE_ALWAYS,
                FILE_ATTRIBUTE_NORMAL,
                nullptr));

            if (!hOut.valid())
            {
                err = L"Failed to create output file.";
                return false;
            }

            const FormatSpec f = GetFormatSpec(fmt.elementType);
            const StyleSpec s = GetStyleSpec(fmt.arrayStyle);

            const size_t elemSize = f.elemSize;
            const size_t elementCount = (elemSize == 0u)
                ? 0u
                : ((fileSize + elemSize - 1u) / elemSize);
            const size_t hexDigits = std::max<size_t>(2u, elemSize * 2u);
            const size_t valuesPerLine = ValuesPerLine(elemSize);

            const uint8_t* data = static_cast<const uint8_t*>(view.get());

            std::string buf;
            buf.reserve(8u * 1024u * 1024u);

            AppendIncludes(f, s, buf);
            AppendHeader(f, s, elementCount, buf);
            if (!WriteAll(hOut, buf.data(), static_cast<DWORD>(buf.size()), err))
                return false;

            buf.clear();

            const DWORD tickStepMs = 120u;
            DWORD lastTick = GetTickCount();

            for (size_t i = 0u; i < elementCount; ++i)
            {
                if ((i % valuesPerLine) == 0u)
                    buf.append("\r\n    ");

                uint64_t value = 0u;
                const size_t base = i * elemSize;
                for (size_t b = 0u; b < elemSize; ++b)
                {
                    const size_t idx = base + b;
                    const uint8_t byte = (idx < fileSize) ? data[idx] : 0u;
                    value |= static_cast<uint64_t>(byte) << (8u * b);
                }

                AppendValueToken(f, value, hexDigits, buf);

                if (i + 1u != elementCount)
                    buf.append(", ");

                if (buf.size() >= (8u * 1024u * 1024u))
                {
                    if (!WriteAll(hOut, buf.data(), static_cast<DWORD>(buf.size()), err))
                        return false;
                    buf.clear();
                }

                const DWORD now = GetTickCount();
                if ((now - lastTick) >= tickStepMs)
                {
                    lastTick = now;
                    const size_t processed = std::min<size_t>(fileSize, (i + 1u) * elemSize);
                    const int pct = (fileSize == 0u) ? 100 : static_cast<int>((processed * 100u) / fileSize);
                    PostMessageW(notifyHwnd, AppMessages::WM_APP_PROGRESS, static_cast<WPARAM>(pct), 0);
                }
            }

            if (!buf.empty())
            {
                if (!WriteAll(hOut, buf.data(), static_cast<DWORD>(buf.size()), err))
                    return false;
                buf.clear();
            }

            AppendFooter(f, s, elementCount, fileSize, buf);
            if (!WriteAll(hOut, buf.data(), static_cast<DWORD>(buf.size()), err))
                return false;

            PostMessageW(notifyHwnd, AppMessages::WM_APP_PROGRESS, 100, 0);
            return true;
        }

        struct WorkerCtx
        {
            Job job{};
            std::wstring* outSmall = nullptr;
        };

        static DWORD WINAPI WorkerThread(LPVOID param)
        {
            WorkerCtx* ctx = static_cast<WorkerCtx*>(param);

            std::wstring err;
            bool ok = false;

            if (ctx->job.largeMode)
            {
                ok = ConvertLargeToFile(ctx->job.inPath, ctx->job.outPath, ctx->job.hwndNotify, ctx->job.format, err);
            }
            else
            {
                std::wstring out;
                ok = ConvertSmallToMemory(ctx->job.inPath, ctx->job.format, out, err);
                if (ok && ctx->outSmall)
                    *(ctx->outSmall) = std::move(out);
            }

            std::wstring msg;
            if (ok)
            {
                if (ctx->job.largeMode)
                    msg = L"OK: saved to file:\r\n" + ctx->job.outPath;
                else
                    msg = L"OK: output generated in UI.";
            }
            else
            {
                msg = L"ERROR:\r\n" + (err.empty() ? L"Conversion failed." : err);
            }

            const size_t bytes = (msg.size() + 1u) * sizeof(wchar_t);
            wchar_t* heapMsg = static_cast<wchar_t*>(HeapAlloc(GetProcessHeap(), 0, bytes));
            if (heapMsg)
                std::memcpy(heapMsg, msg.c_str(), bytes);

            PostMessageW(
                ctx->job.hwndNotify,
                AppMessages::WM_APP_DONE,
                static_cast<WPARAM>(ok ? 1 : 0),
                reinterpret_cast<LPARAM>(heapMsg));

            delete ctx;
            return 0;
        }
    }

    bool GetFileSizeU64(const std::wstring& path, uint64_t& outSize)
    {
        outSize = 0;

        Handle h(CreateFileW(
            path.c_str(),
            GENERIC_READ,
            FILE_SHARE_READ,
            nullptr,
            OPEN_EXISTING,
            FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN,
            nullptr));

        if (!h.valid())
            return false;

        LARGE_INTEGER li{};
        if (!GetFileSizeEx(h, &li))
            return false;

        if (li.QuadPart < 0)
            return false;

        outSize = static_cast<uint64_t>(li.QuadPart);
        return true;
    }

    bool StartConversionAsync(const Job& job, std::wstring& outSmallOutput)
    {
        auto* ctx = new WorkerCtx();
        ctx->job = job;

        if (!job.largeMode)
            ctx->outSmall = &outSmallOutput;
        else
            ctx->outSmall = nullptr;

        Handle hThread(CreateThread(nullptr, 0, WorkerThread, ctx, 0, nullptr));
        if (!hThread.valid())
        {
            delete ctx;
            return false;
        }

        return true;
    }
}
