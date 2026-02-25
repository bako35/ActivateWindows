// ActivateWindows.cpp : Definiuje punkt wejścia dla aplikacji.
//
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Shell32.lib")

#include "framework.h"
#include "windows.h"
#include "windowsx.h"
#include "shlwapi.h"
#include "commctrl.h"
#include "shobjidl.h"
#include "shellapi.h"
#include "resource.h"

#define MAX_LOADSTRING 100

ITaskbarList3* pTaskbar = nullptr;

// Zmienne globalne:
HINSTANCE hInst;                                // bieżące wystąpienie
WCHAR szTitle[MAX_LOADSTRING];                  // Tekst paska tytułu
WCHAR szWindowClass[MAX_LOADSTRING];            // nazwa klasy okna głównego
HFONT hNormalFont = CreateFont(18, 0, 0, 0, FW_DONTCARE, 0, 0, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, OUT_DEFAULT_PRECIS, PROOF_QUALITY, DEFAULT_PITCH, L"Segoe UI"); // czcionka
RECT rc;
SHELLEXECUTEINFO idx01, idx02, idx03, idx11, idx12, idx13, idx21, idx22, idx23, idx31, idx32, idx33, idx41, idx42, idx43, idx51, idx52, idx53, idx61, idx62, idx63, idx71, idx72, idx73, idx81, idx82, idx83, idx91, idx92, idx93;
HWND hwnd, hActiveB, hActivePro, hActiveVerWin, hActiveFrame, hActiveStringWindows, hActiveStringKey1, hActiveStringKey2, hActiveStringKMS1, hActiveStringKMS2, hActiveBCheck, hActiveAbout;

// Przekaż dalej deklaracje funkcji dołączone w tym module kodu:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

BOOL Is_Win10_or_Later() {
    OSVERSIONINFOEX osvi;
    DWORDLONG dwlConditionMask = 0;
    int op = VER_GREATER_EQUAL;

    // Initialize the OSVERSIONINFOEX structure.

    ZeroMemory(&osvi, sizeof(OSVERSIONINFOEX));
    osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);
    osvi.dwMajorVersion = 10;
    osvi.dwMinorVersion = 0;
    osvi.wServicePackMajor = 0;
    osvi.wServicePackMinor = 0;

    // Initialize the condition mask.

    VER_SET_CONDITION(dwlConditionMask, VER_MAJORVERSION, op);
    VER_SET_CONDITION(dwlConditionMask, VER_MINORVERSION, op);
    VER_SET_CONDITION(dwlConditionMask, VER_SERVICEPACKMAJOR, op);
    VER_SET_CONDITION(dwlConditionMask, VER_SERVICEPACKMINOR, op);

    // Perform the test.

    return VerifyVersionInfo(&osvi, VER_MAJORVERSION | VER_MINORVERSION | VER_SERVICEPACKMAJOR | VER_SERVICEPACKMINOR, dwlConditionMask);
}

int APIENTRY wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine, _In_ int nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // TODO: W tym miejscu umieść kod.

    // Inicjuj ciągi globalne
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_ACTIVATEWINDOWS, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    if (!Is_Win10_or_Later()) {
        MessageBox(NULL, L"You must have Windows 10 or Windows 11 to run the program", szTitle, MB_ICONERROR);
        PostQuitMessage(0);
        return 0;
    }

    // Wykonaj inicjowanie aplikacji:
    if (!InitInstance (hInstance, nCmdShow))
    {
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_ACTIVATEWINDOWS));

    MSG msg;

    // Główna pętla komunikatów:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    return (int) msg.wParam;
}



//
//  FUNKCJA: MyRegisterClass()
//
//  PRZEZNACZENIE: Rejestruje klasę okna.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.cbClsExtra     = 0;
    wcex.cbWndExtra     = 0;
    wcex.hInstance      = hInstance;
    wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ACTIVATEWINDOWS));
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_ACTIVATEWINDOWS);
    wcex.lpszClassName  = szWindowClass;
    wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_ACTIVATEWINDOWS));

    return RegisterClassExW(&wcex);
}

//
//   FUNKCJA: InitInstance(HINSTANCE, int)
//
//   PRZEZNACZENIE: Zapisuje dojście wystąpienia i tworzy okno główne
//
//   KOMENTARZE:
//
//        W tej funkcji dojście wystąpienia jest zapisywane w zmiennej globalnej i
//        jest tworzone i wyświetlane okno główne programu.
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   hInst = hInstance; // Przechowuj dojście wystąpienia w naszej zmiennej globalnej

   hwnd = CreateWindowEx(0, szWindowClass, szTitle, WS_SYSMENU|WS_MINIMIZEBOX, 0, 0, 330, 255, NULL, NULL, hInstance, NULL);
   hActiveB = CreateWindowEx(0, L"BUTTON", L"Activate", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON, 108, 180, 100, 30, hwnd, (HMENU)1, hInstance, NULL);
   hActivePro = CreateWindowEx(0, PROGRESS_CLASS, 0, WS_CHILD | WS_VISIBLE, 5, 150, 305, 25, hwnd, 0, hInstance, 0);
   hActiveVerWin = CreateWindowEx(WS_EX_CLIENTEDGE, L"COMBOBOX", NULL, WS_CHILD | WS_VISIBLE | WS_BORDER | CBS_DROPDOWNLIST, 10, 45, 295, 200, hwnd, NULL, hInstance, NULL);
   hActiveFrame = CreateWindowEx(0, L"BUTTON", L"Information", WS_CHILD | WS_VISIBLE | BS_GROUPBOX, 5, 5, 305, 143, hwnd, (HMENU)2, hInstance, NULL);
   hActiveStringWindows = CreateWindowEx(0, L"STATIC", NULL, WS_CHILD | WS_VISIBLE | SS_CENTER, 10, 25, 295, 15, hwnd, NULL, hInstance, NULL);
   hActiveStringKey1 = CreateWindowEx(0, L"STATIC", NULL, WS_CHILD | WS_VISIBLE | SS_CENTER, 10, 75, 295, 15, hwnd, NULL, hInstance, NULL);
   hActiveStringKey2 = CreateWindowEx(0, L"STATIC", NULL, WS_CHILD | WS_VISIBLE | SS_CENTER, 10, 92, 295, 15, hwnd, NULL, hInstance, NULL);
   hActiveStringKMS1 = CreateWindowEx(0, L"STATIC", NULL, WS_CHILD | WS_VISIBLE | SS_CENTER, 10, 110, 295, 15, hwnd, NULL, hInstance, NULL);
   hActiveStringKMS2 = CreateWindowEx(0, L"STATIC", NULL, WS_CHILD | WS_VISIBLE | SS_CENTER, 10, 127, 295, 15, hwnd, NULL, hInstance, NULL);
   hActiveBCheck = CreateWindowEx(0, L"BUTTON", L"Check", WS_CHILD | WS_VISIBLE | WS_TABSTOP, 210, 180, 100, 30, hwnd, (HMENU)3, hInstance, NULL);
   hActiveAbout = CreateWindowEx(0, L"BUTTON", L"About", WS_CHILD | WS_VISIBLE | WS_TABSTOP, 5, 180, 100, 30, hwnd, (HMENU)4, hInstance, NULL);

   SendMessage(hActivePro, PBM_SETRANGE, 0, (LPARAM)MAKELONG(0, 99));
   SendMessage(hActivePro, PBM_SETSTEP, (WPARAM)33, 0);
   SendMessage(hActiveVerWin, CB_ADDSTRING, 0, (LPARAM)L"Windows 10/11 Education");
   SendMessage(hActiveVerWin, CB_ADDSTRING, 0, (LPARAM)L"Windows 10/11 Education N");
   SendMessage(hActiveVerWin, CB_ADDSTRING, 0, (LPARAM)L"Windows 10/11 Enterprise");
   SendMessage(hActiveVerWin, CB_ADDSTRING, 0, (LPARAM)L"Windows 10/11 Enterprise N");
   SendMessage(hActiveVerWin, CB_ADDSTRING, 0, (LPARAM)L"Windows 10/11 Home Country Specific");
   SendMessage(hActiveVerWin, CB_ADDSTRING, 0, (LPARAM)L"Windows 10/11 Home");
   SendMessage(hActiveVerWin, CB_ADDSTRING, 0, (LPARAM)L"Windows 10/11 Home N");
   SendMessage(hActiveVerWin, CB_ADDSTRING, 0, (LPARAM)L"Windows 10/11 Home Single Language");
   SendMessage(hActiveVerWin, CB_ADDSTRING, 0, (LPARAM)L"Windows 10/11 Pro");
   SendMessage(hActiveVerWin, CB_ADDSTRING, 0, (LPARAM)L"Windows 10/11 Pro N");
   SendMessage(hActiveVerWin, WM_SETFONT, (WPARAM)hNormalFont, 0);
   SendMessage(hActiveFrame, WM_SETFONT, (WPARAM)hNormalFont, 0);
   SendMessage(hActiveStringWindows, WM_SETFONT, (WPARAM)hNormalFont, 0);
   SendMessage(hActiveStringKey1, WM_SETFONT, (WPARAM)hNormalFont, 0);
   SendMessage(hActiveStringKey2, WM_SETFONT, (WPARAM)hNormalFont, 0);
   SendMessage(hActiveStringKMS1, WM_SETFONT, (WPARAM)hNormalFont, 0);
   SendMessage(hActiveStringKMS2, WM_SETFONT, (WPARAM)hNormalFont, 0);
   SendMessage(hActiveB, WM_SETFONT, (WPARAM)hNormalFont, 0);
   SendMessage(hActiveBCheck, WM_SETFONT, (WPARAM)hNormalFont, 0);
   SendMessage(hActiveAbout, WM_SETFONT, (WPARAM)hNormalFont, 0);

   SetWindowText(hActiveStringWindows, L"Windows:");
   SetWindowText(hActiveStringKey1, L"Key:");
   SetWindowText(hActiveStringKey2, L"XXXXX-XXXXX-XXXXX-XXXXX-XXXXX");
   SetWindowText(hActiveStringKMS1, L"KMS:");
   SetWindowText(hActiveStringKMS2, L"example.com");

   EnableWindow(hActiveB, FALSE);

   GetWindowRect(hwnd, &rc);
   int xPos = (GetSystemMetrics(SM_CXSCREEN) - rc.right) / 2;
   int yPos = (GetSystemMetrics(SM_CYSCREEN) - rc.bottom) / 2;
   SetWindowPos(hwnd, 0, xPos, yPos, 0, 0, SWP_NOZORDER | SWP_NOSIZE);

   if (!hwnd)
   {
      return FALSE;
   }

   ShowWindow(hwnd, nCmdShow);
   UpdateWindow(hwnd);

   return TRUE;
}

//
//  FUNKCJA: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PRZEZNACZENIE: Przetwarza komunikaty dla okna głównego.
//
//  WM_COMMAND  - przetwarzaj menu aplikacji
//  WM_PAINT    - Maluj okno główne
//  WM_DESTROY  - opublikuj komunikat o wyjściu i wróć
//
//
LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    int index = ComboBox_GetCurSel(hActiveVerWin);
    int msgboxid;
    CoCreateInstance(CLSID_TaskbarList, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pTaskbar));
    switch (message)
    {
    case WM_COMMAND:
        if (LOWORD(wParam) == 4) {
            DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hwnd, About);
            }
        if (LOWORD(wParam) == 3) {
            ShellExecute(NULL, L"open", L"C:\\Windows\\system32\\cmd.exe", L"/c slmgr /dli && slmgr /xpr", NULL, SW_HIDE);
        }
        if ((HWND)lParam == hActiveB && index == 0) {
            EnableWindow(hActiveB, FALSE);
            EnableWindow(hActiveVerWin, FALSE);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 0, 3);
            idx01.cbSize = sizeof(SHELLEXECUTEINFO);
            idx01.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx01.hwnd = NULL;
            idx01.lpVerb = L"runas";
            idx01.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx01.lpParameters = L"/c slmgr /ipk NW6C2-QMPVW-D7KKK-3GKT6-VCFB2";
            idx01.lpDirectory = NULL;
            idx01.nShow = SW_HIDE;
            idx01.hInstApp = NULL;
            ShellExecuteEx(&idx01);
            WaitForSingleObject(idx01.hProcess, INFINITE);
            CloseHandle(idx01.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 1, 3);

            idx02.cbSize = sizeof(SHELLEXECUTEINFO);
            idx02.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx02.hwnd = NULL;
            idx02.lpVerb = L"runas";
            idx02.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx02.lpParameters = L"/c slmgr /skms kms.digiboy.ir";
            idx02.lpDirectory = NULL;
            idx02.nShow = SW_HIDE;
            idx02.hInstApp = NULL;
            ShellExecuteEx(&idx02);
            WaitForSingleObject(idx02.hProcess, INFINITE);
            CloseHandle(idx02.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 2, 3);

            idx03.cbSize = sizeof(SHELLEXECUTEINFO);
            idx03.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx03.hwnd = NULL;
            idx03.lpVerb = L"runas";
            idx03.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx03.lpParameters = L"/c slmgr /ato";
            idx03.lpDirectory = NULL;
            idx03.nShow = SW_HIDE;
            idx03.hInstApp = NULL;
            ShellExecuteEx(&idx03);
            WaitForSingleObject(idx03.hProcess, INFINITE);
            CloseHandle(idx03.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 3, 3);

            if (GetLastError() != 0) {
                pTaskbar->SetProgressState(hwnd, TBPF_ERROR);
                pTaskbar->SetProgressValue(hwnd, 3, 3);
                FlashWindow(hwnd, TRUE);
                MessageBox(hwnd, L"An error occurred while activating Windows", szTitle, MB_ICONERROR | MB_APPLMODAL);
                EnableWindow(hActiveB, TRUE);
                EnableWindow(hActiveVerWin, TRUE);
            }
            else {
                FlashWindow(hwnd, TRUE);
                msgboxid = MessageBox(hwnd, L"Do you want to open activation settings?", szTitle, MB_ICONINFORMATION | MB_APPLMODAL | MB_YESNO);
            }
        }
        if ((HWND)lParam == hActiveVerWin && index == 0) {
            SetWindowText(hActiveStringKey2, L"NW6C2-QMPVW-D7KKK-3GKT6-VCFB2");
            SetWindowText(hActiveStringKMS2, L"kms.digiboy.ir");
            EnableWindow(hActiveB, TRUE);
        }
        if ((HWND)lParam == hActiveB && index == 1) {
            EnableWindow(hActiveB, FALSE);
            EnableWindow(hActiveVerWin, FALSE);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 0, 3);
            idx11.cbSize = sizeof(SHELLEXECUTEINFO);
            idx11.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx11.hwnd = NULL;
            idx11.lpVerb = L"runas";
            idx11.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx11.lpParameters = L"/c slmgr /ipk 2WH4N-8QGBV-H22JP-CT43Q-MDWWJ";
            idx11.lpDirectory = NULL;
            idx11.nShow = SW_HIDE;
            idx11.hInstApp = NULL;
            ShellExecuteEx(&idx11);
            WaitForSingleObject(idx11.hProcess, INFINITE);
            CloseHandle(idx11.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 1, 3);

            idx12.cbSize = sizeof(SHELLEXECUTEINFO);
            idx12.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx12.hwnd = NULL;
            idx12.lpVerb = L"runas";
            idx12.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx12.lpParameters = L"/c slmgr /skms kms.digiboy.ir";
            idx12.lpDirectory = NULL;
            idx12.nShow = SW_HIDE;
            idx12.hInstApp = NULL;
            ShellExecuteEx(&idx12);
            WaitForSingleObject(idx12.hProcess, INFINITE);
            CloseHandle(idx12.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 2, 3);

            idx13.cbSize = sizeof(SHELLEXECUTEINFO);
            idx13.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx13.hwnd = NULL;
            idx13.lpVerb = L"runas";
            idx13.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx13.lpParameters = L"/c slmgr /ato";
            idx13.lpDirectory = NULL;
            idx13.nShow = SW_HIDE;
            idx13.hInstApp = NULL;
            ShellExecuteEx(&idx13);
            WaitForSingleObject(idx13.hProcess, INFINITE);
            CloseHandle(idx13.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 3, 3);

            if (GetLastError() != 0) {
                pTaskbar->SetProgressState(hwnd, TBPF_ERROR);
                pTaskbar->SetProgressValue(hwnd, 3, 3);
                FlashWindow(hwnd, TRUE);
                MessageBox(hwnd, L"An error occurred while activating Windows", szTitle, MB_ICONERROR | MB_APPLMODAL);
                EnableWindow(hActiveB, TRUE);
                EnableWindow(hActiveVerWin, TRUE);
            }
            else {
                FlashWindow(hwnd, TRUE);
                msgboxid = MessageBox(hwnd, L"Do you want to open activation settings?", szTitle, MB_ICONINFORMATION | MB_APPLMODAL | MB_YESNO);
            }
        }
        if ((HWND)lParam == hActiveVerWin && index == 1) {
            SetWindowText(hActiveStringKey2, L"2WH4N-8QGBV-H22JP-CT43Q-MDWWJ");
            SetWindowText(hActiveStringKMS2, L"kms.digiboy.ir");
            EnableWindow(hActiveB, TRUE);
        }
        if ((HWND)lParam == hActiveB && index == 2) {
            EnableWindow(hActiveB, FALSE);
            EnableWindow(hActiveVerWin, FALSE);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 0, 3);
            idx21.cbSize = sizeof(SHELLEXECUTEINFO);
            idx21.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx21.hwnd = NULL;
            idx21.lpVerb = L"runas";
            idx21.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx21.lpParameters = L"/c slmgr /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43";
            idx21.lpDirectory = NULL;
            idx21.nShow = SW_HIDE;
            idx21.hInstApp = NULL;
            ShellExecuteEx(&idx21);
            WaitForSingleObject(idx21.hProcess, INFINITE);
            CloseHandle(idx21.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 1, 3);

            idx22.cbSize = sizeof(SHELLEXECUTEINFO);
            idx22.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx22.hwnd = NULL;
            idx22.lpVerb = L"runas";
            idx22.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx22.lpParameters = L"/c slmgr /skms kms.digiboy.ir";
            idx22.lpDirectory = NULL;
            idx22.nShow = SW_HIDE;
            idx22.hInstApp = NULL;
            ShellExecuteEx(&idx22);
            WaitForSingleObject(idx22.hProcess, INFINITE);
            CloseHandle(idx22.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 2, 3);

            idx23.cbSize = sizeof(SHELLEXECUTEINFO);
            idx23.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx23.hwnd = NULL;
            idx23.lpVerb = L"runas";
            idx23.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx23.lpParameters = L"/c slmgr /ato";
            idx23.lpDirectory = NULL;
            idx23.nShow = SW_HIDE;
            idx23.hInstApp = NULL;
            ShellExecuteEx(&idx23);
            WaitForSingleObject(idx23.hProcess, INFINITE);
            CloseHandle(idx23.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 3, 3);

            if (GetLastError() != 0) {
                pTaskbar->SetProgressState(hwnd, TBPF_ERROR);
                pTaskbar->SetProgressValue(hwnd, 3, 3);
                FlashWindow(hwnd, TRUE);
                MessageBox(hwnd, L"An error occurred while activating Windows", szTitle, MB_ICONERROR | MB_APPLMODAL);
                EnableWindow(hActiveB, TRUE);
                EnableWindow(hActiveVerWin, TRUE);
            }
            else {
                FlashWindow(hwnd, TRUE);
                msgboxid = MessageBox(hwnd, L"Do you want to open activation settings?", szTitle, MB_ICONINFORMATION | MB_APPLMODAL | MB_YESNO);
            }
        }
        if ((HWND)lParam == hActiveVerWin && index == 2) {
            SetWindowText(hActiveStringKey2, L"NPPR9-FWDCX-D2C8J-H872K-2YT43");
            SetWindowText(hActiveStringKMS2, L"kms.digiboy.ir");
            EnableWindow(hActiveB, TRUE);
        }
        if ((HWND)lParam == hActiveB && index == 3) {
            EnableWindow(hActiveB, FALSE);
            EnableWindow(hActiveVerWin, FALSE);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 0, 3);
            idx31.cbSize = sizeof(SHELLEXECUTEINFO);
            idx31.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx31.hwnd = NULL;
            idx31.lpVerb = L"runas";
            idx31.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx31.lpParameters = L"/c slmgr /ipk DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4";
            idx31.lpDirectory = NULL;
            idx31.nShow = SW_HIDE;
            idx31.hInstApp = NULL;
            ShellExecuteEx(&idx31);
            WaitForSingleObject(idx31.hProcess, INFINITE);
            CloseHandle(idx31.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 1, 3);

            idx32.cbSize = sizeof(SHELLEXECUTEINFO);
            idx32.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx32.hwnd = NULL;
            idx32.lpVerb = L"runas";
            idx32.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx32.lpParameters = L"/c slmgr /skms kms.digiboy.ir";
            idx32.lpDirectory = NULL;
            idx32.nShow = SW_HIDE;
            idx32.hInstApp = NULL;
            ShellExecuteEx(&idx32);
            WaitForSingleObject(idx32.hProcess, INFINITE);
            CloseHandle(idx32.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 2, 3);

            idx33.cbSize = sizeof(SHELLEXECUTEINFO);
            idx33.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx33.hwnd = NULL;
            idx33.lpVerb = L"runas";
            idx33.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx33.lpParameters = L"/c slmgr /ato";
            idx33.lpDirectory = NULL;
            idx33.nShow = SW_HIDE;
            idx33.hInstApp = NULL;
            ShellExecuteEx(&idx33);
            WaitForSingleObject(idx33.hProcess, INFINITE);
            CloseHandle(idx33.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 3, 3);

            if (GetLastError() != 0) {
                pTaskbar->SetProgressState(hwnd, TBPF_ERROR);
                pTaskbar->SetProgressValue(hwnd, 3, 3);
                FlashWindow(hwnd, TRUE);
                MessageBox(hwnd, L"An error occurred while activating Windows", szTitle, MB_ICONERROR | MB_APPLMODAL);
                EnableWindow(hActiveB, TRUE);
                EnableWindow(hActiveVerWin, TRUE);
            }
            else {
                FlashWindow(hwnd, TRUE);
                msgboxid = MessageBox(hwnd, L"Do you want to open activation settings?", szTitle, MB_ICONINFORMATION | MB_APPLMODAL | MB_YESNO);
            }
        }
        if ((HWND)lParam == hActiveVerWin && index == 3) {
            SetWindowText(hActiveStringKey2, L"DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4");
            SetWindowText(hActiveStringKMS2, L"kms.digiboy.ir");
            EnableWindow(hActiveB, TRUE);
        }
        if ((HWND)lParam == hActiveB && index == 4) {
            EnableWindow(hActiveB, FALSE);
            EnableWindow(hActiveVerWin, FALSE);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 0, 3);
            idx41.cbSize = sizeof(SHELLEXECUTEINFO);
            idx41.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx41.hwnd = NULL;
            idx41.lpVerb = L"runas";
            idx41.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx41.lpParameters = L"/c slmgr /ipk PVMJN-6DFY6-9CCP6-7BKTT-D3WVR";
            idx41.lpDirectory = NULL;
            idx41.nShow = SW_HIDE;
            idx41.hInstApp = NULL;
            ShellExecuteEx(&idx41);
            WaitForSingleObject(idx41.hProcess, INFINITE);
            CloseHandle(idx41.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 1, 3);

            idx42.cbSize = sizeof(SHELLEXECUTEINFO);
            idx42.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx42.hwnd = NULL;
            idx42.lpVerb = L"runas";
            idx42.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx42.lpParameters = L"/c slmgr /skms kms.digiboy.ir";
            idx42.lpDirectory = NULL;
            idx42.nShow = SW_HIDE;
            idx42.hInstApp = NULL;
            ShellExecuteEx(&idx42);
            WaitForSingleObject(idx42.hProcess, INFINITE);
            CloseHandle(idx42.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 2, 3);

            idx43.cbSize = sizeof(SHELLEXECUTEINFO);
            idx43.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx43.hwnd = NULL;
            idx43.lpVerb = L"runas";
            idx43.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx43.lpParameters = L"/c slmgr /ato";
            idx43.lpDirectory = NULL;
            idx43.nShow = SW_HIDE;
            idx43.hInstApp = NULL;
            ShellExecuteEx(&idx43);
            WaitForSingleObject(idx43.hProcess, INFINITE);
            CloseHandle(idx43.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 3, 3);

            if (GetLastError() != 0) {
                pTaskbar->SetProgressState(hwnd, TBPF_ERROR);
                pTaskbar->SetProgressValue(hwnd, 3, 3);
                FlashWindow(hwnd, TRUE);
                MessageBox(hwnd, L"An error occurred while activating Windows", szTitle, MB_ICONERROR | MB_APPLMODAL);
                EnableWindow(hActiveB, TRUE);
                EnableWindow(hActiveVerWin, TRUE);
            }
            else {
                FlashWindow(hwnd, TRUE);
                msgboxid = MessageBox(hwnd, L"Do you want to open activation settings?", szTitle, MB_ICONINFORMATION | MB_APPLMODAL | MB_YESNO);
            }
        }
        if ((HWND)lParam == hActiveVerWin && index == 4) {
            SetWindowText(hActiveStringKey2, L"PVMJN-6DFY6-9CCP6-7BKTT-D3WVR");
            SetWindowText(hActiveStringKMS2, L"kms.digiboy.ir");
            EnableWindow(hActiveB, TRUE);
        }
        if ((HWND)lParam == hActiveB && index == 5) {
            EnableWindow(hActiveB, FALSE);
            EnableWindow(hActiveVerWin, FALSE);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 0, 3);
            idx51.cbSize = sizeof(SHELLEXECUTEINFO);
            idx51.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx51.hwnd = NULL;
            idx51.lpVerb = L"runas";
            idx51.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx51.lpParameters = L"/c slmgr /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99";
            idx51.lpDirectory = NULL;
            idx51.nShow = SW_HIDE;
            idx51.hInstApp = NULL;
            ShellExecuteEx(&idx51);
            WaitForSingleObject(idx51.hProcess, INFINITE);
            CloseHandle(idx51.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 1, 3);

            idx52.cbSize = sizeof(SHELLEXECUTEINFO);
            idx52.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx52.hwnd = NULL;
            idx52.lpVerb = L"runas";
            idx52.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx52.lpParameters = L"/c slmgr /skms kms.digiboy.ir";
            idx52.lpDirectory = NULL;
            idx52.nShow = SW_HIDE;
            idx52.hInstApp = NULL;
            ShellExecuteEx(&idx52);
            WaitForSingleObject(idx52.hProcess, INFINITE);
            CloseHandle(idx52.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 2, 3);

            idx53.cbSize = sizeof(SHELLEXECUTEINFO);
            idx53.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx53.hwnd = NULL;
            idx53.lpVerb = L"runas";
            idx53.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx53.lpParameters = L"/c slmgr /ato";
            idx53.lpDirectory = NULL;
            idx53.nShow = SW_HIDE;
            idx53.hInstApp = NULL;
            ShellExecuteEx(&idx53);
            WaitForSingleObject(idx53.hProcess, INFINITE);
            CloseHandle(idx53.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 3, 3);

            if (GetLastError() != 0) {
                pTaskbar->SetProgressState(hwnd, TBPF_ERROR);
                pTaskbar->SetProgressValue(hwnd, 3, 3);
                FlashWindow(hwnd, TRUE);
                MessageBox(hwnd, L"An error occurred while activating Windows", szTitle, MB_ICONERROR | MB_APPLMODAL);
                EnableWindow(hActiveB, TRUE);
                EnableWindow(hActiveVerWin, TRUE);
            }
            else {
                FlashWindow(hwnd, TRUE);
                msgboxid = MessageBox(hwnd, L"Do you want to open activation settings?", szTitle, MB_ICONINFORMATION | MB_APPLMODAL | MB_YESNO);
            }
        }
        if ((HWND)lParam == hActiveVerWin && index == 5) {
            SetWindowText(hActiveStringKey2, L"TX9XD-98N7V-6WMQ6-BX7FG-H8Q99");
            SetWindowText(hActiveStringKMS2, L"kms.digiboy.ir");
            EnableWindow(hActiveB, TRUE);
        }
        if ((HWND)lParam == hActiveB && index == 6) {
            EnableWindow(hActiveB, FALSE);
            EnableWindow(hActiveVerWin, FALSE);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 0, 3);
            idx61.cbSize = sizeof(SHELLEXECUTEINFO);
            idx61.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx61.hwnd = NULL;
            idx61.lpVerb = L"runas";
            idx61.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx61.lpParameters = L"/c slmgr /ipk 3KHY7-WNT83-DGQKR-F7HPR-844BM";
            idx61.lpDirectory = NULL;
            idx61.nShow = SW_HIDE;
            idx61.hInstApp = NULL;
            ShellExecuteEx(&idx61);
            WaitForSingleObject(idx61.hProcess, INFINITE);
            CloseHandle(idx61.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 1, 3);

            idx62.cbSize = sizeof(SHELLEXECUTEINFO);
            idx62.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx62.hwnd = NULL;
            idx62.lpVerb = L"runas";
            idx62.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx62.lpParameters = L"/c slmgr /skms kms.digiboy.ir";
            idx62.lpDirectory = NULL;
            idx62.nShow = SW_HIDE;
            idx62.hInstApp = NULL;
            ShellExecuteEx(&idx62);
            WaitForSingleObject(idx62.hProcess, INFINITE);
            CloseHandle(idx62.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 2, 3);

            idx63.cbSize = sizeof(SHELLEXECUTEINFO);
            idx63.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx63.hwnd = NULL;
            idx63.lpVerb = L"runas";
            idx63.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx63.lpParameters = L"/c slmgr /ato";
            idx63.lpDirectory = NULL;
            idx63.nShow = SW_HIDE;
            idx63.hInstApp = NULL;
            ShellExecuteEx(&idx63);
            WaitForSingleObject(idx63.hProcess, INFINITE);
            CloseHandle(idx63.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 3, 3);

            if (GetLastError() != 0) {
                pTaskbar->SetProgressState(hwnd, TBPF_ERROR);
                pTaskbar->SetProgressValue(hwnd, 3, 3);
                FlashWindow(hwnd, TRUE);
                MessageBox(hwnd, L"An error occurred while activating Windows", szTitle, MB_ICONERROR | MB_APPLMODAL);
                EnableWindow(hActiveB, TRUE);
                EnableWindow(hActiveVerWin, TRUE);
            }
            else {
                FlashWindow(hwnd, TRUE);
                msgboxid = MessageBox(hwnd, L"Do you want to open activation settings?", szTitle, MB_ICONINFORMATION | MB_APPLMODAL | MB_YESNO);
            }
        }
        if ((HWND)lParam == hActiveVerWin && index == 6) {
            SetWindowText(hActiveStringKey2, L"3KHY7-WNT83-DGQKR-F7HPR-844BM");
            SetWindowText(hActiveStringKMS2, L"kms.digiboy.ir");
            EnableWindow(hActiveB, TRUE);
        }
        if ((HWND)lParam == hActiveB && index == 7) {
            EnableWindow(hActiveB, FALSE);
            EnableWindow(hActiveVerWin, FALSE);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 0, 3);
            idx71.cbSize = sizeof(SHELLEXECUTEINFO);
            idx71.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx71.hwnd = NULL;
            idx71.lpVerb = L"runas";
            idx71.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx71.lpParameters = L"/c slmgr /ipk 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH";
            idx71.lpDirectory = NULL;
            idx71.nShow = SW_HIDE;
            idx71.hInstApp = NULL;
            ShellExecuteEx(&idx71);
            WaitForSingleObject(idx71.hProcess, INFINITE);
            CloseHandle(idx71.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 1, 3);

            idx72.cbSize = sizeof(SHELLEXECUTEINFO);
            idx72.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx72.hwnd = NULL;
            idx72.lpVerb = L"runas";
            idx72.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx72.lpParameters = L"/c slmgr /skms kms.digiboy.ir";
            idx72.lpDirectory = NULL;
            idx72.nShow = SW_HIDE;
            idx72.hInstApp = NULL;
            ShellExecuteEx(&idx72);
            WaitForSingleObject(idx72.hProcess, INFINITE);
            CloseHandle(idx72.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 2, 3);

            idx73.cbSize = sizeof(SHELLEXECUTEINFO);
            idx73.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx73.hwnd = NULL;
            idx73.lpVerb = L"runas";
            idx73.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx73.lpParameters = L"/c slmgr /ato";
            idx73.lpDirectory = NULL;
            idx73.nShow = SW_HIDE;
            idx73.hInstApp = NULL;
            ShellExecuteEx(&idx73);
            WaitForSingleObject(idx73.hProcess, INFINITE);
            CloseHandle(idx73.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 3, 3);

            if (GetLastError() != 0) {
                pTaskbar->SetProgressState(hwnd, TBPF_ERROR);
                pTaskbar->SetProgressValue(hwnd, 3, 3);
                FlashWindow(hwnd, TRUE);
                MessageBox(hwnd, L"An error occurred while activating Windows", szTitle, MB_ICONERROR | MB_APPLMODAL);
                EnableWindow(hActiveB, TRUE);
                EnableWindow(hActiveVerWin, TRUE);
            }
            else {
                FlashWindow(hwnd, TRUE);
                msgboxid = MessageBox(hwnd, L"Do you want to open activation settings?", szTitle, MB_ICONINFORMATION | MB_APPLMODAL | MB_YESNO);
            }
        }
        if ((HWND)lParam == hActiveVerWin && index == 7) {
            SetWindowText(hActiveStringKey2, L"7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH");
            SetWindowText(hActiveStringKMS2, L"kms.digiboy.ir");
            EnableWindow(hActiveB, TRUE);
        }
        if ((HWND)lParam == hActiveB && index == 8) {
            EnableWindow(hActiveB, FALSE);
            EnableWindow(hActiveVerWin, FALSE);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 0, 3);
            idx81.cbSize = sizeof(SHELLEXECUTEINFO);
            idx81.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx81.hwnd = NULL;
            idx81.lpVerb = L"runas";
            idx81.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx81.lpParameters = L"/c slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX";
            idx81.lpDirectory = NULL;
            idx81.nShow = SW_HIDE;
            idx81.hInstApp = NULL;
            ShellExecuteEx(&idx81);
            WaitForSingleObject(idx81.hProcess, INFINITE);
            CloseHandle(idx81.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 1, 3);

            idx82.cbSize = sizeof(SHELLEXECUTEINFO);
            idx82.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx82.hwnd = NULL;
            idx82.lpVerb = L"runas";
            idx82.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx82.lpParameters = L"/c slmgr /skms kms.digiboy.ir";
            idx82.lpDirectory = NULL;
            idx82.nShow = SW_HIDE;
            idx82.hInstApp = NULL;
            ShellExecuteEx(&idx82);
            WaitForSingleObject(idx82.hProcess, INFINITE);
            CloseHandle(idx82.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 2, 3);

            idx83.cbSize = sizeof(SHELLEXECUTEINFO);
            idx83.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx83.hwnd = NULL;
            idx83.lpVerb = L"runas";
            idx83.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx83.lpParameters = L"/c slmgr /ato";
            idx83.lpDirectory = NULL;
            idx83.nShow = SW_HIDE;
            idx83.hInstApp = NULL;
            ShellExecuteEx(&idx83);
            WaitForSingleObject(idx83.hProcess, INFINITE);
            CloseHandle(idx83.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 3, 3);

            if (GetLastError() != 0) {
                pTaskbar->SetProgressState(hwnd, TBPF_ERROR);
                pTaskbar->SetProgressValue(hwnd, 3, 3);
                FlashWindow(hwnd, TRUE);
                MessageBox(hwnd, L"An error occurred while activating Windows", szTitle, MB_ICONERROR | MB_APPLMODAL);
                EnableWindow(hActiveB, TRUE);
                EnableWindow(hActiveVerWin, TRUE);
            }
            else {
                FlashWindow(hwnd, TRUE);
                msgboxid = MessageBox(hwnd, L"Do you want to open activation settings?", szTitle, MB_ICONINFORMATION | MB_APPLMODAL | MB_YESNO);
            }
        }
        if ((HWND)lParam == hActiveVerWin && index == 8) {
            SetWindowText(hActiveStringKey2, L"W269N-WFGWX-YVC9B-4J6C9-T83GX");
            SetWindowText(hActiveStringKMS2, L"kms.digiboy.ir");
            EnableWindow(hActiveB, TRUE);
        }
        if ((HWND)lParam == hActiveB && index == 9) {
            EnableWindow(hActiveB, FALSE);
            EnableWindow(hActiveVerWin, FALSE);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 0, 3);
            idx91.cbSize = sizeof(SHELLEXECUTEINFO);
            idx91.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx91.hwnd = NULL;
            idx91.lpVerb = L"runas";
            idx91.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx91.lpParameters = L"/c slmgr /ipk MH37W-N47XK-V7XM9-C7227-GCQG9";
            idx91.lpDirectory = NULL;
            idx91.nShow = SW_HIDE;
            idx91.hInstApp = NULL;
            ShellExecuteEx(&idx91);
            WaitForSingleObject(idx91.hProcess, INFINITE);
            CloseHandle(idx91.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 1, 3);

            idx92.cbSize = sizeof(SHELLEXECUTEINFO);
            idx92.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx92.hwnd = NULL;
            idx92.lpVerb = L"runas";
            idx92.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx92.lpParameters = L"/c slmgr /skms kms.digiboy.ir";
            idx92.lpDirectory = NULL;
            idx92.nShow = SW_HIDE;
            idx92.hInstApp = NULL;
            ShellExecuteEx(&idx92);
            WaitForSingleObject(idx92.hProcess, INFINITE);
            CloseHandle(idx92.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 2, 3);

            idx93.cbSize = sizeof(SHELLEXECUTEINFO);
            idx93.fMask = SEE_MASK_NOCLOSEPROCESS;
            idx93.hwnd = NULL;
            idx93.lpVerb = L"runas";
            idx93.lpFile = L"C:\\Windows\\system32\\cmd.exe";
            idx93.lpParameters = L"/c slmgr /ato";
            idx93.lpDirectory = NULL;
            idx93.nShow = SW_HIDE;
            idx93.hInstApp = NULL;
            ShellExecuteEx(&idx93);
            WaitForSingleObject(idx93.hProcess, INFINITE);
            CloseHandle(idx93.hProcess);
            SendMessage(hActivePro, PBM_STEPIT, 0, 0);
            pTaskbar->SetProgressState(hwnd, TBPF_NORMAL);
            pTaskbar->SetProgressValue(hwnd, 3, 3);

            if (GetLastError() != 0) {
                pTaskbar->SetProgressState(hwnd, TBPF_ERROR);
                pTaskbar->SetProgressValue(hwnd, 3, 3);
                FlashWindow(hwnd, TRUE);
                MessageBox(hwnd, L"An error occurred while activating Windows", szTitle, MB_ICONERROR | MB_APPLMODAL);
                EnableWindow(hActiveB, TRUE);
                EnableWindow(hActiveVerWin, TRUE);
            }
            else {
                FlashWindow(hwnd, TRUE);
                msgboxid = MessageBox(hwnd, L"Do you want to open activation settings?", szTitle, MB_ICONINFORMATION | MB_APPLMODAL | MB_YESNO);
            }
        }
        if ((HWND)lParam == hActiveVerWin && index == 9) {
            SetWindowText(hActiveStringKey2, L"MH37W-N47XK-V7XM9-C7227-GCQG9");
            SetWindowText(hActiveStringKMS2, L"kms.digiboy.ir");
            EnableWindow(hActiveB, TRUE);
        }
        switch (msgboxid)
        {
        case IDYES:
            ShellExecute(NULL, L"open", L"ms-settings:activation", NULL, NULL, SW_NORMAL);
            DestroyWindow(hwnd);
            break;
        case IDNO:
            DestroyWindow(hwnd);
            break;
        }
        break;
    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hwnd, &ps);
            // TODO: Tutaj dodaj kod rysujący używający elementu hdc...
            EndPaint(hwnd, &ps);
        }
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hwnd, message, wParam, lParam);
    }
    return 0;
}

// Procedura obsługi komunikatów dla okna informacji o programie.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        if (LOWORD(wParam) == IDYES) {
            ShellExecute(NULL, L"open", L"http://github.com/bako35", NULL, NULL, SW_NORMAL);
            EndDialog(hDlg, LOWORD(wParam));
        }
        break;
    }
    return (INT_PTR)FALSE;
}
