// CoreServices.cpp
#include "CoreServices.h"

#include <commdlg.h>

#include <cstring>
#include <cwchar>
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

        static constexpr wchar_t HEXW[16] = {
            L'0', L'1', L'2', L'3', L'4', L'5', L'6', L'7', L'8', L'9', L'A', L'B', L'C', L'D', L'E', L'F'
        };

        static constexpr char HEXA[16] = {
            '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F'
        };

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

        static void AppendHexByte(std::string& dst, uint8_t b, bool withCommaSpace)
        {
            dst.push_back('0');
            dst.push_back('x');
            dst.push_back(HEXA[(b >> 4) & 0x0F]);
            dst.push_back(HEXA[b & 0x0F]);
            if (withCommaSpace)
            {
                dst.push_back(',');
                dst.push_back(' ');
            }
        }

        static void BuildCArrayWideFast(const uint8_t* data, size_t n, std::wstring& out)
        {
            static constexpr wchar_t HEADER[]  = L"const unsigned char fileBytes[] = {";
            static constexpr wchar_t INDENT[]  = L"\r\n    ";
            static constexpr wchar_t FOOTER1[] = L"\r\n};\r\n";
            static constexpr wchar_t FOOTER2[] = L"const size_t fileBytesSize = sizeof(fileBytes);\r\n";

            constexpr size_t headerLen  = (sizeof(HEADER)  / sizeof(wchar_t)) - 1u;
            constexpr size_t indentLen  = (sizeof(INDENT)  / sizeof(wchar_t)) - 1u;
            constexpr size_t footer1Len = (sizeof(FOOTER1) / sizeof(wchar_t)) - 1u;
            constexpr size_t footer2Len = (sizeof(FOOTER2) / sizeof(wchar_t)) - 1u;

            const size_t lines = (n == 0u) ? 0u : ((n + 15u) / 16u);
            const size_t sepCount = (n > 0u) ? (n - 1u) : 0u;

            const size_t totalLen =
                headerLen +
                (lines * indentLen) +
                (n * 4u) +
                (sepCount * 2u) +
                footer1Len +
                footer2Len;

            out.clear();
            out.resize(totalLen);

            wchar_t* p = out.data();

            std::memcpy(p, HEADER, headerLen * sizeof(wchar_t));
            p += headerLen;

            for (size_t i = 0u; i < n; ++i)
            {
                if ((i % 16u) == 0u)
                {
                    std::memcpy(p, INDENT, indentLen * sizeof(wchar_t));
                    p += indentLen;
                }

                const uint8_t b = data[i];

                *p++ = L'0';
                *p++ = L'x';
                *p++ = HEXW[(b >> 4) & 0x0Fu];
                *p++ = HEXW[b & 0x0Fu];

                if (i + 1u != n)
                {
                    *p++ = L',';
                    *p++ = L' ';
                }
            }

            std::memcpy(p, FOOTER1, footer1Len * sizeof(wchar_t));
            p += footer1Len;

            std::memcpy(p, FOOTER2, footer2Len * sizeof(wchar_t));
            p += footer2Len;
        }

        static bool ConvertSmallToMemory(const std::wstring& path, std::wstring& out, std::wstring& err)
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
            BuildCArrayWideFast(data, fileSize, out);
            return true;
        }

        static bool ConvertLargeToFile(
            const std::wstring& inPath,
            const std::wstring& outPath,
            HWND notifyHwnd,
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

            const char* header1 = "const unsigned char fileBytes[] = {";
            const char* footer1 = "\r\n};\r\n";
            const char* footer2 = "const size_t fileBytesSize = sizeof(fileBytes);\r\n";

            if (!WriteAll(hOut, header1, static_cast<DWORD>(std::strlen(header1)), err))
                return false;

            const uint8_t* data = static_cast<const uint8_t*>(view.get());

            std::string buf;
            buf.reserve(8u * 1024u * 1024u);

            const DWORD tickStepMs = 120u;
            DWORD lastTick = GetTickCount();

            for (size_t i = 0u; i < fileSize; ++i)
            {
                if ((i % 16u) == 0u)
                    buf.append("\r\n    ");

                const bool withComma = (i + 1u != fileSize);
                AppendHexByte(buf, data[i], withComma);

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
                    const int pct = (fileSize == 0u) ? 100 : static_cast<int>((i * 100u) / fileSize);
                    PostMessageW(notifyHwnd, AppMessages::WM_APP_PROGRESS, static_cast<WPARAM>(pct), 0);
                }
            }

            if (!buf.empty())
            {
                if (!WriteAll(hOut, buf.data(), static_cast<DWORD>(buf.size()), err))
                    return false;
            }

            if (!WriteAll(hOut, footer1, static_cast<DWORD>(std::strlen(footer1)), err))
                return false;

            if (!WriteAll(hOut, footer2, static_cast<DWORD>(std::strlen(footer2)), err))
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
                ok = ConvertLargeToFile(ctx->job.inPath, ctx->job.outPath, ctx->job.hwndNotify, err);
            }
            else
            {
                std::wstring out;
                ok = ConvertSmallToMemory(ctx->job.inPath, out, err);
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
