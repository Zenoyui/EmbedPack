// App.cpp
#include "App.h"
#include "CoreServices.h"

#include <windows.h>
#include <commctrl.h>
#include <richedit.h>

#include <cstdint>
#include <cwchar>
#include <string>
#include <utility>
#include <algorithm>

#pragma comment(lib, "Comctl32.lib")

namespace EmbedPack
{
    namespace
    {
        int DpiGetForWindowSafe(HWND hwnd)
        {
            using Fn = UINT(WINAPI*)(HWND);
            static Fn pGetDpiForWindow = reinterpret_cast<Fn>(
                GetProcAddress(GetModuleHandleW(L"user32.dll"), "GetDpiForWindow"));

            if (pGetDpiForWindow && hwnd)
                return (int)pGetDpiForWindow(hwnd);

            HDC dc = GetDC(hwnd);
            const int dpi = dc ? GetDeviceCaps(dc, LOGPIXELSX) : 96;
            if (dc) ReleaseDC(hwnd, dc);
            return dpi;
        }

        int DpiScale(int v, int dpi) { return (v * dpi + 48) / 96; }

        void EnablePerMonitorDpiAware()
        {
            using Fn = BOOL(WINAPI*)(DPI_AWARENESS_CONTEXT);
            static Fn pSet = reinterpret_cast<Fn>(
                GetProcAddress(GetModuleHandleW(L"user32.dll"), "SetProcessDpiAwarenessContext"));

            if (pSet)
                pSet(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
        }

        
        struct Theme
        {
            COLORREF bg          = RGB(30, 30, 30);
            COLORREF panel       = RGB(37, 37, 38);
            COLORREF panel2      = RGB(45, 45, 48);
            COLORREF border      = RGB(60, 60, 60);
            COLORREF borderSoft  = RGB(50, 50, 50);

            COLORREF text        = RGB(212, 212, 212);
            COLORREF textDim     = RGB(160, 160, 160);

            COLORREF accent      = RGB(0, 122, 204);
            COLORREF good        = RGB(64, 185, 120);
            COLORREF bad         = RGB(235, 92, 92);

            COLORREF btnIdle     = RGB(45, 45, 48);
            COLORREF btnHover    = RGB(62, 62, 64);
            COLORREF btnDown     = RGB(80, 80, 84);
            COLORREF btnDisabled = RGB(35, 35, 36);

            COLORREF editBg      = RGB(30, 30, 30);
            COLORREF editText    = RGB(212, 212, 212);
        };

        Theme g_theme{};

        void FillRectColor(HDC dc, const RECT& rc, COLORREF c)
        {
            HBRUSH b = CreateSolidBrush(c);
            FillRect(dc, &rc, b);
            DeleteObject(b);
        }

        void FrameRectColor(HDC dc, const RECT& rc, COLORREF c)
        {
            HPEN pen = CreatePen(PS_SOLID, 1, c);
            HGDIOBJ oldPen = SelectObject(dc, pen);
            HGDIOBJ oldBrush = SelectObject(dc, GetStockObject(HOLLOW_BRUSH));
            Rectangle(dc, rc.left, rc.top, rc.right, rc.bottom);
            SelectObject(dc, oldBrush);
            SelectObject(dc, oldPen);
            DeleteObject(pen);
        }

        void DrawTextEllipsized(HDC dc, const wchar_t* text, RECT rc, UINT format, COLORREF color)
        {
            SetTextColor(dc, color);
            SetBkMode(dc, TRANSPARENT);
            DrawTextW(dc, text, -1, &rc, format | DT_END_ELLIPSIS);
        }

        static int MeasureTextPx(HDC dc, const std::wstring& s)
        {
            if (s.empty())
                return 0;

            SIZE sz{};
            GetTextExtentPoint32W(dc, s.c_str(), (int)s.size(), &sz);
            return (int)sz.cx;
        }

        static std::wstring NormalizeSlashes(std::wstring s)
        {
            for (auto& ch : s)
            {
                if (ch == L'/') ch = L'\\';
            }
            return s;
        }

        static std::wstring PathEllipsizeMiddle(HDC dc, const std::wstring& path, int maxWidthPx)
        {
            std::wstring p = NormalizeSlashes(path);

            if (p.empty() || maxWidthPx <= 0)
                return L"";

            if (MeasureTextPx(dc, p) <= maxWidthPx)
                return p;

            const std::wstring ell = L"...\\";
            if (MeasureTextPx(dc, ell) >= maxWidthPx)
            {
                const std::wstring ell2 = L"...";
                if (MeasureTextPx(dc, ell2) <= maxWidthPx) return ell2;
                return L"";
            }

            size_t lastSlash = p.find_last_of(L'\\');
            std::wstring filePart;
            std::wstring headPart;

            if (lastSlash != std::wstring::npos && lastSlash + 1 < p.size())
            {
                filePart = p.substr(lastSlash + 1);
                headPart = p.substr(0, lastSlash);
            }
            else
            {
                filePart = p;
                headPart.clear();
            }

            std::wstring suffix = filePart;
            if (suffix.empty())
                suffix = p;

            if (MeasureTextPx(dc, suffix) > maxWidthPx)
            {
                std::wstring s = suffix;
                const std::wstring dots = L"...";
                int lo = 0;
                int hi = (int)s.size();
                std::wstring best = dots;

                while (lo <= hi)
                {
                    int mid = (lo + hi) / 2;
                    std::wstring cand = s.substr(0, mid) + dots;
                    if (MeasureTextPx(dc, cand) <= maxWidthPx)
                    {
                        best = cand;
                        lo = mid + 1;
                    }
                    else
                    {
                        hi = mid - 1;
                    }
                }
                return best;
            }

            int suffixW = MeasureTextPx(dc, suffix);
            int ellW = MeasureTextPx(dc, ell);
            int remaining = maxWidthPx - (ellW + suffixW);

            if (remaining <= 0)
            {
                return ell + suffix;
            }

            std::wstring prefixCandidate;

            if (p.size() >= 3 && p[1] == L':' && p[2] == L'\\')
            {
                prefixCandidate = p.substr(0, 3);
            }
            else if (p.size() >= 2 && p[0] == L'\\' && p[1] == L'\\')
            {
                size_t s1 = p.find(L'\\', 2);
                if (s1 != std::wstring::npos)
                {
                    size_t s2 = p.find(L'\\', s1 + 1);
                    if (s2 != std::wstring::npos)
                    {
                        size_t s3 = p.find(L'\\', s2 + 1);
                        if (s3 != std::wstring::npos)
                            prefixCandidate = p.substr(0, s3 + 1);
                        else
                            prefixCandidate = p + L"\\";
                    }
                    else
                    {
                        prefixCandidate = p + L"\\";
                    }
                }
                else
                {
                    prefixCandidate = L"\\\\";
                }
            }
            else
            {
                size_t s0 = p.find(L'\\');
                if (s0 != std::wstring::npos)
                    prefixCandidate = p.substr(0, s0 + 1);
                else
                    prefixCandidate = L"";
            }

            if (MeasureTextPx(dc, prefixCandidate) > remaining)
            {
                std::wstring pref = prefixCandidate;
                int lo = 0;
                int hi = (int)pref.size();
                std::wstring best;

                while (lo <= hi)
                {
                    int mid = (lo + hi) / 2;
                    std::wstring cand = pref.substr(0, mid);
                    if (MeasureTextPx(dc, cand) <= remaining)
                    {
                        best = cand;
                        lo = mid + 1;
                    }
                    else
                    {
                        hi = mid - 1;
                    }
                }

                return best + ell + suffix;
            }

            int usedPrefix = MeasureTextPx(dc, prefixCandidate);
            int extra = remaining - usedPrefix;

            std::wstring prefix = prefixCandidate;

            size_t maxTake = (lastSlash != std::wstring::npos) ? lastSlash + 1 : p.size();
            if (prefix.size() < maxTake && extra > 0)
            {
                size_t base = prefix.size();
                size_t lo = base;
                size_t hi = maxTake;
                size_t bestLen = base;

                while (lo <= hi)
                {
                    size_t mid = (lo + hi) / 2;
                    std::wstring cand = p.substr(0, mid);
                    if (MeasureTextPx(dc, cand) <= usedPrefix + extra)
                    {
                        bestLen = mid;
                        lo = mid + 1;
                    }
                    else
                    {
                        if (mid == 0) break;
                        hi = mid - 1;
                    }
                }

                prefix = p.substr(0, bestLen);
            }

            return prefix + ell + suffix;
        }

        struct ButtonState
        {
            bool hot = false;
            bool down = false;
        };

        constexpr int ID_BTN_SELECT  = 1001;
        constexpr int ID_BTN_CONVERT = 1002;
        constexpr int ID_BTN_COPY    = 1003;

        constexpr int ID_TT_PATH = 2001;

        struct Layout
        {
            int pad = 12;
            int toolbarH = 54;
            int statusH  = 28;

            int btnH = 30;
            int btnW1 = 124;
            int btnW2 = 108;
            int btnW3 = 92;

            int gap = 10;

            int editPad = 12;
        };

        constexpr int MIN_CLIENT_W_96 = 600;
        constexpr int MIN_CLIENT_H_96 = 400;

        bool AdjustWindowRectExForDpiSafe(RECT* rc, DWORD style, BOOL hasMenu, DWORD exStyle, UINT dpi)
        {
            using Fn = BOOL(WINAPI*)(LPRECT, DWORD, BOOL, DWORD, UINT);
            static Fn pAdjustForDpi = reinterpret_cast<Fn>(
                GetProcAddress(GetModuleHandleW(L"user32.dll"), "AdjustWindowRectExForDpi"));

            if (pAdjustForDpi)
                return pAdjustForDpi(rc, style, hasMenu, exStyle, dpi);

            return AdjustWindowRectEx(rc, style, hasMenu, exStyle);
        }

        POINT ComputeMinTrackSize(HWND hwnd)
        {
            const int dpi = DpiGetForWindowSafe(hwnd);
            const int clientMinW = DpiScale(MIN_CLIENT_W_96, dpi);
            const int clientMinH = DpiScale(MIN_CLIENT_H_96, dpi);

            RECT rc{ 0, 0, clientMinW, clientMinH };

            const DWORD style = (DWORD)GetWindowLongPtrW(hwnd, GWL_STYLE);
            const DWORD exStyle = (DWORD)GetWindowLongPtrW(hwnd, GWL_EXSTYLE);

            AdjustWindowRectExForDpiSafe(&rc, style, FALSE, exStyle, (UINT)dpi);

            POINT pt{};
            pt.x = rc.right - rc.left;
            pt.y = rc.bottom - rc.top;
            return pt;
        }

        struct StreamCookie
        {
            const BYTE* bytes = nullptr;
            size_t sizeBytes = 0;
            size_t posBytes = 0;
        };

        DWORD CALLBACK RichEditStreamInCallback(DWORD_PTR cookie, LPBYTE buffer, LONG cb, LONG* pcb)
        {
            if (!pcb) return 1;

            StreamCookie* sc = reinterpret_cast<StreamCookie*>(cookie);
            if (!sc || !sc->bytes || sc->posBytes > sc->sizeBytes)
            {
                *pcb = 0;
                return 0;
            }

            const size_t remaining = sc->sizeBytes - sc->posBytes;
            const size_t toCopy = std::min<size_t>((size_t)cb, remaining);

            if (toCopy > 0)
                memcpy(buffer, sc->bytes + sc->posBytes, toCopy);

            sc->posBytes += toCopy;
            *pcb = (LONG)toCopy;
            return 0;
        }

        class UiWindow final
        {
        public:
            explicit UiWindow(HINSTANCE hInstance);
            bool CreateAndShow(int nCmdShow);
            HWND Hwnd() const noexcept { return m_hwnd; }

        private:
            friend LRESULT CALLBACK ButtonSubProc(HWND, UINT, WPARAM, LPARAM, UINT_PTR, DWORD_PTR);

            static LRESULT CALLBACK WndProcSetup(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
            static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
            LRESULT HandleMessage(UINT msg, WPARAM wParam, LPARAM lParam);

            void OnCreate();
            void OnDestroy();
            void OnSize(int w, int h);
            void OnPaint();
            void OnDrawItem(const DRAWITEMSTRUCT* dis);
            HBRUSH OnCtlColorEdit(HDC dc, HWND hCtrl);
            HBRUSH OnCtlColorStatic(HDC dc, HWND hCtrl);

            void OnSelectFile();
            void OnConvert();
            void OnCopy();
            void OnProgress(int pct);
            void OnDone(bool ok, wchar_t* heapMsg);

            void LockUi(bool lock);
            void UpdateStatusText(const std::wstring& s);
            void UpdatePathText(const std::wstring& s);
            void UpdatePathTooltip();
            void SetOutputText(const std::wstring& s);
            void SetBusyCursor(bool busy);
            void InvalidateToolbarAndStatus();

            void RecomputeDpi();
            void ApplyFonts();
            void LayoutChildren(int clientW, int clientH);

            void TrackHotButton(HWND btn);
            void SetButtonHot(HWND btn, bool hot);
            void SetButtonDown(HWND btn, bool down);

            ButtonState* GetStateFor(HWND btn);

            void CacheEditRect();

            void ConfigureRichEditAppearance();

        private:
            HINSTANCE m_hInstance = nullptr;
            HWND m_hwnd = nullptr;

            HWND m_btnSelect = nullptr;
            HWND m_btnConvert = nullptr;
            HWND m_btnCopy = nullptr;

            HWND m_lblPath = nullptr;
            HWND m_editOutput = nullptr;

            HWND m_ttPath = nullptr;

            HFONT m_fontUi = nullptr;
            HFONT m_fontMono = nullptr;

            HBRUSH m_brBg = nullptr;
            HBRUSH m_brPanel = nullptr;
            HBRUSH m_brEdit = nullptr;

            Layout m_lay{};
            int m_dpi = 96;

            std::wstring m_selectedFilePath;
            std::wstring m_outputW;
            std::wstring m_statusText = L"Ready";
            std::wstring m_pathText = L"No input file selected";

            int m_progress = 0;
            bool m_busy = false;

            ButtonState m_stateSelect{};
            ButtonState m_stateConvert{};
            ButtonState m_stateCopy{};

            bool m_uiLocked = false;
            bool m_lastOk = true;

            RECT m_rcEditClient{};

            HMODULE m_hMsftEdit = nullptr;
        };

        UiWindow::UiWindow(HINSTANCE hInstance)
            : m_hInstance(hInstance)
        {
            SetRectEmpty(&m_rcEditClient);
        }

        void UiWindow::RecomputeDpi()
        {
            m_dpi = DpiGetForWindowSafe(m_hwnd);

            m_lay.pad      = DpiScale(12, m_dpi);
            m_lay.toolbarH = DpiScale(54, m_dpi);
            m_lay.statusH  = DpiScale(28, m_dpi);

            m_lay.btnH  = DpiScale(30, m_dpi);
            m_lay.btnW1 = DpiScale(124, m_dpi);
            m_lay.btnW2 = DpiScale(108, m_dpi);
            m_lay.btnW3 = DpiScale(92,  m_dpi);

            m_lay.gap = DpiScale(10, m_dpi);
            m_lay.editPad = DpiScale(12, m_dpi);

            if (m_fontUi) { DeleteObject(m_fontUi); m_fontUi = nullptr; }
            if (m_fontMono) { DeleteObject(m_fontMono); m_fontMono = nullptr; }

            LOGFONTW lf{};
            lf.lfHeight = -DpiScale(14, m_dpi);
            lf.lfWeight = FW_NORMAL;
            wcscpy_s(lf.lfFaceName, L"Segoe UI");
            m_fontUi = CreateFontIndirectW(&lf);

            LOGFONTW lm{};
            lm.lfHeight = -DpiScale(13, m_dpi);
            lm.lfWeight = FW_NORMAL;
            wcscpy_s(lm.lfFaceName, L"Consolas");
            m_fontMono = CreateFontIndirectW(&lm);

            ApplyFonts();
            ConfigureRichEditAppearance();
        }

        void UiWindow::ApplyFonts()
        {
            if (m_btnSelect)  SendMessageW(m_btnSelect,  WM_SETFONT, (WPARAM)m_fontUi, TRUE);
            if (m_btnConvert) SendMessageW(m_btnConvert, WM_SETFONT, (WPARAM)m_fontUi, TRUE);
            if (m_btnCopy)    SendMessageW(m_btnCopy,    WM_SETFONT, (WPARAM)m_fontUi, TRUE);

            if (m_lblPath)    SendMessageW(m_lblPath,    WM_SETFONT, (WPARAM)m_fontUi, TRUE);
            if (m_editOutput) SendMessageW(m_editOutput, WM_SETFONT, (WPARAM)m_fontMono, TRUE);
        }

        void UiWindow::ConfigureRichEditAppearance()
        {
            if (!m_editOutput)
                return;

            SendMessageW(m_editOutput, EM_SETBKGNDCOLOR, 0, (LPARAM)g_theme.editBg);

            CHARFORMAT2W cf{};
            cf.cbSize = sizeof(cf);
            cf.dwMask = CFM_COLOR | CFM_FACE | CFM_SIZE;
            cf.crTextColor = g_theme.editText;
            wcscpy_s(cf.szFaceName, L"Consolas");
            cf.yHeight = 13 * 20;

            SendMessageW(m_editOutput, EM_SETCHARFORMAT, (WPARAM)SCF_ALL, (LPARAM)&cf);

            SendMessageW(m_editOutput, EM_SETREADONLY, TRUE, 0);

            SendMessageW(m_editOutput, EM_EXLIMITTEXT, 0, (LPARAM)0x7FFFFFFF);

            SendMessageW(m_editOutput, EM_SETTARGETDEVICE, 0, 0);

            const int tabTwips = 4 * 1440 / 10;
            SendMessageW(m_editOutput, EM_SETTABSTOPS, 1, (LPARAM)&tabTwips);
        }

        void UiWindow::CacheEditRect()
        {
            if (!m_editOutput)
            {
                SetRectEmpty(&m_rcEditClient);
                return;
            }

            RECT r{};
            GetClientRect(m_editOutput, &r);

            POINT tl{ r.left, r.top };
            POINT br{ r.right, r.bottom };
            ClientToScreen(m_editOutput, &tl);
            ClientToScreen(m_editOutput, &br);

            ScreenToClient(m_hwnd, &tl);
            ScreenToClient(m_hwnd, &br);

            m_rcEditClient = { tl.x, tl.y, br.x, br.y };
        }

        void UiWindow::InvalidateToolbarAndStatus()
        {
            RECT rc{};
            GetClientRect(m_hwnd, &rc);

            RECT top{ rc.left, rc.top, rc.right, rc.top + m_lay.toolbarH };
            RECT bottom{ rc.left, rc.bottom - m_lay.statusH, rc.right, rc.bottom };

            InvalidateRect(m_hwnd, &top, FALSE);
            InvalidateRect(m_hwnd, &bottom, FALSE);
        }

        void UiWindow::SetBusyCursor(bool busy)
        {
            m_busy = busy;
            if (busy)
                SetCursor(LoadCursorW(nullptr, IDC_WAIT));
            else
                SetCursor(LoadCursorW(nullptr, IDC_ARROW));
        }

        void UiWindow::UpdateStatusText(const std::wstring& s)
        {
            m_statusText = s;
            InvalidateToolbarAndStatus();
        }

        void UiWindow::UpdatePathTooltip()
        {
            if (!m_ttPath || !m_lblPath)
                return;

            TOOLINFOW ti{};
            ti.cbSize = sizeof(ti);
            ti.hwnd = m_hwnd;
            ti.uId = (UINT_PTR)m_lblPath;
            ti.uFlags = TTF_IDISHWND | TTF_SUBCLASS;
            ti.lpszText = const_cast<wchar_t*>(m_pathText.c_str());

            SendMessageW(m_ttPath, TTM_UPDATETIPTEXTW, 0, (LPARAM)&ti);
        }

        void UiWindow::UpdatePathText(const std::wstring& s)
        {
            m_pathText = s;

            if (m_lblPath)
            {
                SetWindowTextW(m_lblPath, m_pathText.c_str());
            }

            UpdatePathTooltip();
            InvalidateToolbarAndStatus();
        }

        void UiWindow::SetOutputText(const std::wstring& s)
        {
            if (!m_editOutput)
                return;

            SendMessageW(m_editOutput, WM_SETREDRAW, FALSE, 0);

            SendMessageW(m_editOutput, EM_SETSEL, 0, -1);
            SendMessageW(m_editOutput, EM_REPLACESEL, FALSE, (LPARAM)L"");

            StreamCookie sc{};
            sc.bytes = reinterpret_cast<const BYTE*>(s.data());
            sc.sizeBytes = s.size() * sizeof(wchar_t);
            sc.posBytes = 0;

            EDITSTREAM es{};
            es.dwCookie = (DWORD_PTR)&sc;
            es.dwError = 0;
            es.pfnCallback = RichEditStreamInCallback;

            SendMessageW(m_editOutput, EM_STREAMIN, (WPARAM)(SF_TEXT | SF_UNICODE), (LPARAM)&es);

            SendMessageW(m_editOutput, EM_SETSEL, 0, 0);
            SendMessageW(m_editOutput, EM_SCROLLCARET, 0, 0);

            SendMessageW(m_editOutput, WM_SETREDRAW, TRUE, 0);
            InvalidateRect(m_editOutput, nullptr, TRUE);
        }

        void UiWindow::TrackHotButton(HWND btn)
        {
            TRACKMOUSEEVENT tme{};
            tme.cbSize = sizeof(tme);
            tme.dwFlags = TME_LEAVE;
            tme.hwndTrack = btn;
            TrackMouseEvent(&tme);
        }

        ButtonState* UiWindow::GetStateFor(HWND btn)
        {
            if (btn == m_btnSelect)  return &m_stateSelect;
            if (btn == m_btnConvert) return &m_stateConvert;
            if (btn == m_btnCopy)    return &m_stateCopy;
            return nullptr;
        }

        void UiWindow::SetButtonHot(HWND btn, bool hot)
        {
            ButtonState* st = GetStateFor(btn);
            if (!st) return;
            if (st->hot == hot) return;
            st->hot = hot;
            InvalidateRect(btn, nullptr, TRUE);
        }

        void UiWindow::SetButtonDown(HWND btn, bool down)
        {
            ButtonState* st = GetStateFor(btn);
            if (!st) return;
            if (st->down == down) return;
            st->down = down;
            InvalidateRect(btn, nullptr, TRUE);
        }

        void UiWindow::LayoutChildren(int clientW, int clientH)
        {
            const int pad = m_lay.pad;

            const int toolbarH = m_lay.toolbarH;
            const int statusH  = m_lay.statusH;

            const int innerW = clientW - (pad * 2);
            const int innerH = clientH - toolbarH - statusH - (pad * 2);

            int x = pad;
            int y = pad;

            const int btnH = m_lay.btnH;
            const int gap  = m_lay.gap;

            const int yBtn = y + (toolbarH - btnH) / 2;

            int xBtn = x;

            if (m_btnSelect)
                SetWindowPos(m_btnSelect, nullptr, xBtn, yBtn, m_lay.btnW1, btnH, SWP_NOZORDER | SWP_NOACTIVATE);
            xBtn += m_lay.btnW1 + gap;

            if (m_btnConvert)
                SetWindowPos(m_btnConvert, nullptr, xBtn, yBtn, m_lay.btnW2, btnH, SWP_NOZORDER | SWP_NOACTIVATE);
            xBtn += m_lay.btnW2 + gap;

            if (m_btnCopy)
                SetWindowPos(m_btnCopy, nullptr, xBtn, yBtn, m_lay.btnW3, btnH, SWP_NOZORDER | SWP_NOACTIVATE);

            if (m_lblPath)
            {
                const int lblX = xBtn + m_lay.btnW3 + gap;
                const int lblW = (pad + innerW) - lblX;

                SetWindowPos(m_lblPath, nullptr, lblX, yBtn, lblW, btnH, SWP_NOZORDER | SWP_NOACTIVATE);
                InvalidateRect(m_lblPath, nullptr, TRUE);
            }

            const int editY = y + toolbarH + pad;
            const int editH = innerH - pad;
            const int editX = x;
            const int editW = innerW;

            if (m_editOutput)
                SetWindowPos(m_editOutput, nullptr, editX, editY, editW, editH, SWP_NOZORDER | SWP_NOACTIVATE);

            CacheEditRect();
        }

        void UiWindow::OnSize(int w, int h)
        {
            LayoutChildren(w, h);
            InvalidateRect(m_hwnd, nullptr, TRUE);
        }

        HBRUSH UiWindow::OnCtlColorEdit(HDC dc, HWND)
        {
            SetBkColor(dc, g_theme.editBg);
            SetTextColor(dc, g_theme.editText);
            return m_brEdit;
        }

        HBRUSH UiWindow::OnCtlColorStatic(HDC dc, HWND hCtrl)
        {
            SetBkMode(dc, TRANSPARENT);

            if (hCtrl == m_lblPath)
                SetTextColor(dc, g_theme.textDim);
            else
                SetTextColor(dc, g_theme.textDim);

            return m_brPanel;
        }

        void UiWindow::OnDrawItem(const DRAWITEMSTRUCT* dis)
        {
            if (!dis)
                return;

            if (dis->CtlType != ODT_BUTTON)
                return;

            HWND btn = dis->hwndItem;
            ButtonState* st = GetStateFor(btn);

            const bool disabled = (dis->itemState & ODS_DISABLED) != 0;
            const bool pressed  = (dis->itemState & ODS_SELECTED) != 0;

            COLORREF bg = g_theme.btnIdle;
            if (disabled)
                bg = g_theme.btnDisabled;
            else if (pressed || (st && st->down))
                bg = g_theme.btnDown;
            else if (st && st->hot)
                bg = g_theme.btnHover;

            FillRectColor(dis->hDC, dis->rcItem, bg);

            RECT rc = dis->rcItem;
            FrameRectColor(dis->hDC, rc, g_theme.borderSoft);

            wchar_t text[128]{};
            GetWindowTextW(btn, text, (int)std::size(text));

            SetBkMode(dis->hDC, TRANSPARENT);
            SetTextColor(dis->hDC, disabled ? g_theme.textDim : g_theme.text);

            RECT trc = dis->rcItem;
            trc.left += DpiScale(8, m_dpi);
            trc.right -= DpiScale(8, m_dpi);

            DrawTextW(dis->hDC, text, -1, &trc, DT_SINGLELINE | DT_VCENTER | DT_CENTER);
        }

        void UiWindow::OnPaint()
        {
            PAINTSTRUCT ps{};
            HDC dc = BeginPaint(m_hwnd, &ps);

            RECT rc{};
            GetClientRect(m_hwnd, &rc);

            FillRectColor(dc, rc, g_theme.bg);

            RECT top{ rc.left, rc.top, rc.right, rc.top + m_lay.toolbarH };
            FillRectColor(dc, top, g_theme.panel);

            RECT bottom{ rc.left, rc.bottom - m_lay.statusH, rc.right, rc.bottom };
            FillRectColor(dc, bottom, g_theme.panel);

            FrameRectColor(dc, top, g_theme.borderSoft);
            FrameRectColor(dc, bottom, g_theme.borderSoft);

            RECT sRc = bottom;
            sRc.left += m_lay.pad;
            sRc.right -= m_lay.pad;

            DrawTextEllipsized(dc, m_statusText.c_str(), sRc, DT_SINGLELINE | DT_VCENTER | DT_LEFT, g_theme.textDim);

            if (!IsRectEmpty(&m_rcEditClient))
            {
                RECT fr = m_rcEditClient;
                InflateRect(&fr, 1, 1);
                FrameRectColor(dc, fr, g_theme.border);
            }

            EndPaint(m_hwnd, &ps);
        }

        void UiWindow::OnCreate()
        {
            m_brBg    = CreateSolidBrush(g_theme.bg);
            m_brPanel = CreateSolidBrush(g_theme.panel);
            m_brEdit  = CreateSolidBrush(g_theme.editBg);

            m_hMsftEdit = LoadLibraryW(L"Msftedit.dll");

            RecomputeDpi();

            m_btnSelect = CreateWindowExW(
                0, L"BUTTON", L"Select File",
                WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
                0, 0, 0, 0,
                m_hwnd, (HMENU)(INT_PTR)ID_BTN_SELECT, m_hInstance, nullptr);

            m_btnConvert = CreateWindowExW(
                0, L"BUTTON", L"Convert",
                WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
                0, 0, 0, 0,
                m_hwnd, (HMENU)ID_BTN_CONVERT, m_hInstance, nullptr);

            m_btnCopy = CreateWindowExW(
                0, L"BUTTON", L"Copy",
                WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
                0, 0, 0, 0,
                m_hwnd, (HMENU)ID_BTN_COPY, m_hInstance, nullptr);

            EnableWindow(m_btnCopy, FALSE);

            m_lblPath = CreateWindowExW(
                0, L"STATIC", L"",
                WS_CHILD | WS_VISIBLE | SS_LEFT | SS_CENTERIMAGE | SS_PATHELLIPSIS | SS_NOPREFIX,
                0, 0, 0, 0,
                m_hwnd, nullptr, m_hInstance, nullptr);

            m_editOutput = CreateWindowExW(
                WS_EX_CLIENTEDGE, L"RICHEDIT50W", L"",
                WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY,
                0, 0, 0, 0,
                m_hwnd, nullptr, m_hInstance, nullptr);

            // Tooltip for full (non-ellipsized) path
            m_ttPath = CreateWindowExW(
                WS_EX_TOPMOST,
                TOOLTIPS_CLASSW,
                nullptr,
                WS_POPUP | TTS_NOPREFIX | TTS_ALWAYSTIP,
                CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
                m_hwnd,
                (HMENU)(INT_PTR)ID_TT_PATH,
                m_hInstance,
                nullptr);

            if (m_ttPath && m_lblPath)
            {
                SetWindowPos(m_ttPath, HWND_TOPMOST, 0, 0, 0, 0,
                             SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);

                TOOLINFOW ti{};
                ti.cbSize = sizeof(ti);
                ti.hwnd = m_hwnd;
                ti.uId = (UINT_PTR)m_lblPath; // using HWND as ID
                ti.uFlags = TTF_IDISHWND | TTF_SUBCLASS;
                ti.lpszText = const_cast<wchar_t*>(m_pathText.c_str());
                SendMessageW(m_ttPath, TTM_ADDTOOLW, 0, (LPARAM)&ti);

                SendMessageW(m_ttPath, TTM_SETMAXTIPWIDTH, 0, (LPARAM)DpiScale(900, m_dpi));
            }

            ApplyFonts();
            ConfigureRichEditAppearance();

            UpdatePathText(L"No input file selected");
            UpdateStatusText(L"Ready");
            SetOutputText(L"Select a file to begin.\r\nThen click Convert.");

            RECT rc{};
            GetClientRect(m_hwnd, &rc);
            LayoutChildren(rc.right - rc.left, rc.bottom - rc.top);
        }

        void UiWindow::OnDestroy()
        {
            if (m_fontUi) { DeleteObject(m_fontUi); m_fontUi = nullptr; }
            if (m_fontMono) { DeleteObject(m_fontMono); m_fontMono = nullptr; }

            if (m_brBg) { DeleteObject(m_brBg); m_brBg = nullptr; }
            if (m_brPanel) { DeleteObject(m_brPanel); m_brPanel = nullptr; }
            if (m_brEdit) { DeleteObject(m_brEdit); m_brEdit = nullptr; }

            if (m_ttPath) { DestroyWindow(m_ttPath); m_ttPath = nullptr; }

            if (m_hMsftEdit) { FreeLibrary(m_hMsftEdit); m_hMsftEdit = nullptr; }
        }

        bool UiWindow::CreateAndShow(int nCmdShow)
        {
            WNDCLASSEXW wc{};
            wc.cbSize = sizeof(wc);
            wc.lpfnWndProc = UiWindow::WndProcSetup;
            wc.hInstance = m_hInstance;
            wc.lpszClassName = L"EmbedPackWindowClass";
            wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
            wc.hbrBackground = nullptr;
            wc.style = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;

            if (!RegisterClassExW(&wc))
                return false;

            const int w = 900;
            const int h = 680;

            m_hwnd = CreateWindowExW(
                0,
                wc.lpszClassName,
                L"EmbedPack Converter",
                WS_OVERLAPPEDWINDOW,
                CW_USEDEFAULT, CW_USEDEFAULT,
                w, h,
                nullptr,
                nullptr,
                m_hInstance,
                this);

            if (!m_hwnd)
                return false;

            ShowWindow(m_hwnd, nCmdShow);
            UpdateWindow(m_hwnd);
            return true;
        }

        LRESULT CALLBACK UiWindow::WndProcSetup(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
        {
            if (msg == WM_NCCREATE)
            {
                auto* cs = reinterpret_cast<CREATESTRUCTW*>(lParam);
                auto* self = reinterpret_cast<UiWindow*>(cs->lpCreateParams);

                SetWindowLongPtrW(hwnd, GWLP_USERDATA, (LONG_PTR)self);
                SetWindowLongPtrW(hwnd, GWLP_WNDPROC, (LONG_PTR)UiWindow::WndProcThunk);

                self->m_hwnd = hwnd;
                return self->HandleMessage(msg, wParam, lParam);
            }

            return DefWindowProcW(hwnd, msg, wParam, lParam);
        }

        LRESULT CALLBACK UiWindow::WndProcThunk(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
        {
            auto* self = reinterpret_cast<UiWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
            if (!self)
                return DefWindowProcW(hwnd, msg, wParam, lParam);

            self->m_hwnd = hwnd;
            return self->HandleMessage(msg, wParam, lParam);
        }

        void UiWindow::OnSelectFile()
        {
            std::wstring path;
            if (!FileDialogs::PromptOpenInputFile(m_hwnd, path))
                return;

            m_selectedFilePath = std::move(path);
            m_outputW.clear();
            EnableWindow(m_btnCopy, FALSE);

            UpdatePathText(m_selectedFilePath);
            UpdateStatusText(L"Ready");
            m_progress = 0;
            m_lastOk = true;

            SetOutputText(L"File selected.\r\nClick Convert to generate output.");
        }

        void UiWindow::LockUi(bool lock)
        {
            m_uiLocked = lock;

            EnableWindow(m_btnSelect,  lock ? FALSE : TRUE);
            EnableWindow(m_btnConvert, lock ? FALSE : TRUE);

            const bool canCopy = (!lock && !m_outputW.empty());
            EnableWindow(m_btnCopy, canCopy ? TRUE : FALSE);

            InvalidateRect(m_btnSelect, nullptr, TRUE);
            InvalidateRect(m_btnConvert, nullptr, TRUE);
            InvalidateRect(m_btnCopy, nullptr, TRUE);
        }

        void UiWindow::OnConvert()
        {
            if (m_selectedFilePath.empty())
            {
                MessageBoxW(m_hwnd, L"Please select a file first.", L"Error", MB_OK | MB_ICONERROR);
                return;
            }

            uint64_t fsize = 0;
            if (!Converter::GetFileSizeU64(m_selectedFilePath, fsize))
            {
                MessageBoxW(m_hwnd, L"Failed to query file size.", L"Error", MB_OK | MB_ICONERROR);
                return;
            }

            const bool largeMode = (fsize > Converter::UI_SOFT_LIMIT);

            std::wstring outPath;
            if (largeMode)
            {
                if (!FileDialogs::PromptSaveOutputPath(m_hwnd, m_selectedFilePath, outPath))
                    return;
            }

            LockUi(true);
            SetBusyCursor(true);
            m_outputW.clear();
            m_progress = 0;
            m_lastOk = true;

            if (largeMode)
            {
                UpdateStatusText(L"Converting (large file mode: saving to disk) ...");
                SetOutputText(L"Converting large file.\r\nOutput will be saved to the selected file.\r\n\r\nProgress: 0%");
            }
            else
            {
                UpdateStatusText(L"Converting ...");
                SetOutputText(L"Converting ...\r\n\r\nProgress: 0%");
            }

            Converter::Job job{};
            job.hwndNotify = m_hwnd;
            job.inPath = m_selectedFilePath;
            job.outPath = outPath;
            job.largeMode = largeMode;

            if (!Converter::StartConversionAsync(job, m_outputW))
            {
                SetBusyCursor(false);
                LockUi(false);
                MessageBoxW(m_hwnd, L"Failed to start worker thread.", L"Error", MB_OK | MB_ICONERROR);
                return;
            }

            InvalidateToolbarAndStatus();
        }

        void UiWindow::OnCopy()
        {
            if (m_outputW.empty())
            {
                MessageBoxW(m_hwnd, L"No data to copy.", L"Error", MB_OK | MB_ICONERROR);
                return;
            }

            Clipboard::SetClipboardUnicode(m_hwnd, m_outputW);
            UpdateStatusText(L"Copied to clipboard");
            MessageBoxW(m_hwnd, L"Copied to clipboard.", L"Success", MB_OK | MB_ICONINFORMATION);
        }

        void UiWindow::OnProgress(int pct)
        {
            m_progress = pct;
            if (pct < 0) m_progress = 0;
            if (pct > 100) m_progress = 100;

            wchar_t buf[128]{};
            swprintf_s(buf, L"Converting ... %d%%", m_progress);
            UpdateStatusText(buf);

            wchar_t t[64]{};
            swprintf_s(t, L"\r\n\r\nProgress: %d%%", m_progress);

            if (m_progress < 100)
            {
                std::wstring s = L"Converting ...";
                s += t;
                SetOutputText(s);
            }

            InvalidateToolbarAndStatus();
        }

        void UiWindow::OnDone(bool ok, wchar_t* heapMsg)
        {
            SetBusyCursor(false);
            LockUi(false);

            m_lastOk = ok;
            m_progress = 100;

            if (ok)
            {
                UpdateStatusText(L"Done");
                if (!m_outputW.empty())
                {
                    SetOutputText(m_outputW);
                    EnableWindow(m_btnCopy, TRUE);
                }
                else
                {
                    if (heapMsg)
                        SetOutputText(heapMsg);
                    EnableWindow(m_btnCopy, FALSE);
                }
            }
            else
            {
                UpdateStatusText(L"Error");
                if (heapMsg)
                    SetOutputText(heapMsg);
                EnableWindow(m_btnCopy, FALSE);
            }

            if (heapMsg)
                HeapFree(GetProcessHeap(), 0, heapMsg);

            InvalidateToolbarAndStatus();
        }

        LRESULT UiWindow::HandleMessage(UINT msg, WPARAM wParam, LPARAM lParam)
        {
            switch (msg)
            {
            case WM_CREATE:
                OnCreate();
                return 0;

            case WM_DESTROY:
                OnDestroy();
                PostQuitMessage(0);
                return 0;

            case WM_SIZE:
                OnSize(LOWORD(lParam), HIWORD(lParam));
                return 0;

            case WM_GETMINMAXINFO:
            {
                auto* mmi = reinterpret_cast<MINMAXINFO*>(lParam);
                const POINT ptMin = ComputeMinTrackSize(m_hwnd);
                mmi->ptMinTrackSize.x = ptMin.x;
                mmi->ptMinTrackSize.y = ptMin.y;
                return 0;
            }

            case WM_PAINT:
                OnPaint();
                return 0;

            case WM_DRAWITEM:
                OnDrawItem(reinterpret_cast<const DRAWITEMSTRUCT*>(lParam));
                return TRUE;

            case WM_COMMAND:
            {
                switch (LOWORD(wParam))
                {
                case ID_BTN_SELECT:  OnSelectFile(); return 0;
                case ID_BTN_CONVERT: OnConvert();    return 0;
                case ID_BTN_COPY:    OnCopy();       return 0;
                default: break;
                }
                return 0;
            }

            case WM_CTLCOLORSTATIC:
                return reinterpret_cast<LRESULT>(OnCtlColorStatic(reinterpret_cast<HDC>(wParam), (HWND)lParam));

            case WM_CTLCOLOREDIT:
                return reinterpret_cast<LRESULT>(OnCtlColorEdit(reinterpret_cast<HDC>(wParam), (HWND)lParam));

            case WM_SETCURSOR:
                if (m_busy)
                {
                    SetCursor(LoadCursorW(nullptr, IDC_WAIT));
                    return TRUE;
                }
                break;

            case WM_DPICHANGED:
            {
                const RECT* r = reinterpret_cast<const RECT*>(lParam);
                SetWindowPos(m_hwnd, nullptr, r->left, r->top, r->right - r->left, r->bottom - r->top,
                             SWP_NOZORDER | SWP_NOACTIVATE);

                RecomputeDpi();

                if (m_ttPath)
                    SendMessageW(m_ttPath, TTM_SETMAXTIPWIDTH, 0, (LPARAM)DpiScale(900, m_dpi));

                RECT rc{};
                GetClientRect(m_hwnd, &rc);
                LayoutChildren(rc.right - rc.left, rc.bottom - rc.top);
                InvalidateRect(m_hwnd, nullptr, TRUE);
                return 0;
            }

            case AppMessages::WM_APP_PROGRESS:
                OnProgress((int)wParam);
                return 0;

            case AppMessages::WM_APP_DONE:
                OnDone(wParam == 1, reinterpret_cast<wchar_t*>(lParam));
                return 0;

            default:
                break;
            }

            return DefWindowProcW(m_hwnd, msg, wParam, lParam);
        }

        LRESULT CALLBACK ButtonSubProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                       UINT_PTR, DWORD_PTR refData)
        {
            UiWindow* self = reinterpret_cast<UiWindow*>(refData);
            if (!self) return DefSubclassProc(hwnd, msg, wParam, lParam);

            switch (msg)
            {
            case WM_MOUSEMOVE:
                self->SetButtonHot(hwnd, true);
                self->TrackHotButton(hwnd);
                return DefSubclassProc(hwnd, msg, wParam, lParam);

            case WM_MOUSELEAVE:
                self->SetButtonHot(hwnd, false);
                self->SetButtonDown(hwnd, false);
                return DefSubclassProc(hwnd, msg, wParam, lParam);

            case WM_LBUTTONDOWN:
                self->SetButtonDown(hwnd, true);
                return DefSubclassProc(hwnd, msg, wParam, lParam);

            case WM_LBUTTONUP:
                self->SetButtonDown(hwnd, false);
                return DefSubclassProc(hwnd, msg, wParam, lParam);

            default:
                break;
            }

            return DefSubclassProc(hwnd, msg, wParam, lParam);
        }
    } 

    App::App(HINSTANCE hInstance)
        : m_hInstance(hInstance)
    {
    }

    int App::Run(int nCmdShow)
    {
        EnablePerMonitorDpiAware();

        INITCOMMONCONTROLSEX icc{};
        icc.dwSize = sizeof(icc);
        icc.dwICC = ICC_STANDARD_CLASSES;
        InitCommonControlsEx(&icc);

        UiWindow window(m_hInstance);
        if (!window.CreateAndShow(nCmdShow))
            return 0;

        HWND hwndMain = window.Hwnd();
        HWND b1 = GetDlgItem(hwndMain, ID_BTN_SELECT);
        HWND b2 = GetDlgItem(hwndMain, ID_BTN_CONVERT);
        HWND b3 = GetDlgItem(hwndMain, ID_BTN_COPY);

        if (b1) SetWindowSubclass(b1, ButtonSubProc, 1, (DWORD_PTR)&window);
        if (b2) SetWindowSubclass(b2, ButtonSubProc, 2, (DWORD_PTR)&window);
        if (b3) SetWindowSubclass(b3, ButtonSubProc, 3, (DWORD_PTR)&window);

        MSG msg{};
        while (GetMessageW(&msg, nullptr, 0, 0) > 0)
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
        return (int)msg.wParam;
    }
}
