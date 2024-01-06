#include <windows.h>
#include <windowsx.h>
#include <shlwapi.h>
#include <commctrl.h>

LPSTR NazwaKlasy = "ActiveWindowsX32";
MSG Komunikat;
HWND hwnd, hActiveB, hActivePro, hActiveVerWin, hActiveFrame, hActiveStringWindows, hActiveStringKey1, hActiveStringKMS1, hActiveStringKey2, hActiveStringKMS2, hActiveBCheck, hActiveAbout;
//HFONT hNormalFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
HFONT hNormalFont = CreateFont(18, 0, 0, 0, FW_DONTCARE, 0, 0, 0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, OUT_DEFAULT_PRECIS, PROOF_QUALITY, DEFAULT_PITCH, "Segoe UI");
RECT rc;
SHELLEXECUTEINFO idx01, idx02, idx03, idx11, idx12, idx13, idx21, idx22, idx23, idx31, idx32, idx33, idx41, idx42, idx43, idx51, idx52, idx53, idx61, idx62, idx63, idx71, idx72, idx73, idx81, idx82, idx83, idx91, idx92, idx93;

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);

BOOL Is_Win10_or_Later(){
   OSVERSIONINFOEX osvi;
   DWORDLONG dwlConditionMask = 0;
   int op=VER_GREATER_EQUAL;

   // Initialize the OSVERSIONINFOEX structure.

   ZeroMemory(&osvi, sizeof(OSVERSIONINFOEX));
   osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);
   osvi.dwMajorVersion = 6;
   osvi.dwMinorVersion = 2;
   osvi.wServicePackMajor = 0;
   osvi.wServicePackMinor = 0;

   // Initialize the condition mask.

   VER_SET_CONDITION(dwlConditionMask, VER_MAJORVERSION, op);
   VER_SET_CONDITION(dwlConditionMask, VER_MINORVERSION, op);
   VER_SET_CONDITION(dwlConditionMask, VER_SERVICEPACKMAJOR, op);
   VER_SET_CONDITION(dwlConditionMask, VER_SERVICEPACKMINOR, op);

   // Perform the test.

   return VerifyVersionInfo(&osvi, VER_MAJORVERSION|VER_MINORVERSION|VER_SERVICEPACKMAJOR|VER_SERVICEPACKMINOR, dwlConditionMask);
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
   
    // WYPE£NIANIE STRUKTURY
    WNDCLASSEX wc;
   
    wc.cbSize = sizeof(WNDCLASSEX);
    wc.style = 0;
    wc.lpfnWndProc = WndProc;
    wc.cbClsExtra = 0;
    wc.cbWndExtra = 0;
    wc.hInstance = hInstance;
    wc.hIcon = LoadIcon(hInstance, "A");
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground =(HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszMenuName = NULL;
    wc.lpszClassName = NazwaKlasy;
    wc.hIconSm = LoadIcon(hInstance, "A");
   
    // REJESTROWANIE KLASY OKNA
    if(!RegisterClassEx(&wc))
    {
        MessageBox(NULL, "Error RegisterClassEx", "Activate Windows 10/11 x32", MB_ICONEXCLAMATION);
        PostQuitMessage(0);
    	return 0;
    }
    
    if(!Is_Win10_or_Later()){
    	MessageBox(NULL, "You must have Windows 10 or Windows 11 to run the program", "Activate Windows 10/11 x32", MB_ICONERROR);
		PostQuitMessage(0);
    	return 0;
	}
   
    // TWORZENIE OKNA
    hwnd = CreateWindowEx(0, NazwaKlasy, "Activate Windows 10/11 x32", WS_SYSMENU, 0, 0, 320, 245, NULL, NULL, hInstance, NULL);
    hActiveB = CreateWindowEx(0, "BUTTON", "Activate", WS_CHILD|WS_VISIBLE|WS_TABSTOP|BS_DEFPUSHBUTTON, 108, 180, 100, 30, hwnd, NULL, hInstance, NULL);
    hActivePro = CreateWindowEx(0, PROGRESS_CLASS, 0, WS_CHILD|WS_VISIBLE, 5, 150, 305, 25, hwnd, 0, hInstance, 0);
    hActiveVerWin = CreateWindowEx(WS_EX_CLIENTEDGE, "COMBOBOX", NULL, WS_CHILD|WS_VISIBLE|WS_BORDER|CBS_DROPDOWNLIST, 10, 45, 295, 200, hwnd, NULL, hInstance, NULL);
    hActiveFrame = CreateWindowEx(0, "BUTTON", "Information", WS_CHILD|WS_VISIBLE|BS_GROUPBOX, 5, 5, 305, 143, hwnd, NULL, hInstance, NULL);
    hActiveStringWindows = CreateWindowEx(0, "STATIC", NULL, WS_CHILD|WS_VISIBLE|SS_CENTER, 10, 25, 295, 15, hwnd, NULL, hInstance, NULL);
    hActiveStringKey1 = CreateWindowEx(0, "STATIC", NULL, WS_CHILD|WS_VISIBLE|SS_CENTER, 10, 75, 295, 15, hwnd, NULL, hInstance, NULL);
    hActiveStringKey2 = CreateWindowEx(0, "STATIC", NULL, WS_CHILD|WS_VISIBLE|SS_CENTER, 10, 92, 295, 15, hwnd, NULL, hInstance, NULL);
    hActiveStringKMS1 = CreateWindowEx(0, "STATIC", NULL, WS_CHILD|WS_VISIBLE|SS_CENTER, 10, 110, 295, 15, hwnd, NULL, hInstance, NULL);
    hActiveStringKMS2 = CreateWindowEx(0, "STATIC", NULL, WS_CHILD|WS_VISIBLE|SS_CENTER, 10, 127, 295, 15, hwnd, NULL, hInstance, NULL);
    hActiveBCheck = CreateWindowEx(0, "BUTTON", "Check", WS_CHILD|WS_VISIBLE|WS_TABSTOP, 210, 180, 100, 30, hwnd, NULL, hInstance, NULL);
    hActiveAbout = CreateWindowEx(0, "BUTTON", "About", WS_CHILD|WS_VISIBLE|WS_TABSTOP, 5, 180, 100, 30, hwnd, NULL, hInstance, NULL);
    
    SendMessage(hActivePro, PBM_SETRANGE, 0, (LPARAM)MAKELONG(0, 99));
    SendMessage(hActivePro, PBM_SETSTEP, (WPARAM)33, 0);
    SendMessage(hActiveVerWin, CB_ADDSTRING, 0, (LPARAM)"Windows 10/11 Education");
    SendMessage(hActiveVerWin, CB_ADDSTRING, 0, (LPARAM)"Windows 10/11 Education N");
    SendMessage(hActiveVerWin, CB_ADDSTRING, 0, (LPARAM)"Windows 10/11 Enterprise");
    SendMessage(hActiveVerWin, CB_ADDSTRING, 0, (LPARAM)"Windows 10/11 Enterprise N");
    SendMessage(hActiveVerWin, CB_ADDSTRING, 0, (LPARAM)"Windows 10/11 Home Country Specific");
    SendMessage(hActiveVerWin, CB_ADDSTRING, 0, (LPARAM)"Windows 10/11 Home");
    SendMessage(hActiveVerWin, CB_ADDSTRING, 0, (LPARAM)"Windows 10/11 Home N");
    SendMessage(hActiveVerWin, CB_ADDSTRING, 0, (LPARAM)"Windows 10/11 Home Single Language");
    SendMessage(hActiveVerWin, CB_ADDSTRING, 0, (LPARAM)"Windows 10/11 Pro");
    SendMessage(hActiveVerWin, CB_ADDSTRING, 0, (LPARAM)"Windows 10/11 Pro N");
    SendMessage(hActiveVerWin, WM_SETFONT, (WPARAM)hNormalFont, 0);
    SendMessage(hActiveFrame, WM_SETFONT, (WPARAM)hNormalFont, 0);
    SendMessage(hActiveStringWindows, WM_SETFONT, (WPARAM)hNormalFont, 0);
    SendMessage(hActiveStringKey1, WM_SETFONT, (WPARAM)hNormalFont, 0);
    //SendMessage(hActiveStringKey2, WM_SETFONT, (WPARAM)hNormalFont, 0);
    SendMessage(hActiveStringKMS1, WM_SETFONT, (WPARAM)hNormalFont, 0);
    //SendMessage(hActiveStringKMS2, WM_SETFONT, (WPARAM)hNormalFont, 0);
    SendMessage(hActiveB, WM_SETFONT, (WPARAM)hNormalFont, 0);
    SendMessage(hActiveBCheck, WM_SETFONT, (WPARAM)hNormalFont, 0);
    SendMessage(hActiveAbout, WM_SETFONT, (WPARAM)hNormalFont, 0);
    
    SetWindowText(hActiveStringWindows, "Windows:");
    SetWindowText(hActiveStringKey1, "Key:");
    SetWindowText(hActiveStringKey2, "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX");
    SetWindowText(hActiveStringKMS1, "KMS:");
    SetWindowText(hActiveStringKMS2, "example.com");
    
    EnableWindow(hActiveB, FALSE);
   
   	GetWindowRect(hwnd, &rc);
   	int xPos = (GetSystemMetrics(SM_CXSCREEN) - rc.right)/2;
   	int yPos = (GetSystemMetrics(SM_CYSCREEN) - rc.bottom)/2;
   	SetWindowPos(hwnd, 0, xPos, yPos, 0, 0, SWP_NOZORDER|SWP_NOSIZE);
   
    if(hwnd == NULL)
    {
        MessageBox(NULL, "hwnd == NULL", "Activate Windows 10/11 x32", MB_ICONEXCLAMATION);
        PostQuitMessage(0);
    	return 0;
    }
   
   HANDLE hMutex = CreateMutex(NULL, FALSE, "AW10/11");
	if(hMutex && GetLastError() != 0){
    	MessageBox(NULL, "This application is already running", "Activate Windows 10/11 x32", MB_ICONEXCLAMATION);
    	PostQuitMessage(0);
    	return 0;
}
   
    ShowWindow(hwnd, nCmdShow); // Poka¿ okienko...
    UpdateWindow(hwnd);
   
    // Pêtla komunikatów
    while( GetMessage(& Komunikat, NULL, 0, 0))
    {
        TranslateMessage(& Komunikat);
        DispatchMessage(& Komunikat);
    }
    return Komunikat.wParam;
}

// OBS£UGA ZDARZEÑ
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	int index = ComboBox_GetCurSel(hActiveVerWin);
	int msgboxid;
    switch(msg)
    {
    case WM_CLOSE:
        DestroyWindow(hwnd);
        break;
       
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
       
    case WM_COMMAND:
    	if((HWND) lParam == hActiveB && index == 0){
    		EnableWindow(hActiveB, FALSE);
    		EnableWindow(hActiveVerWin, FALSE);
    		idx01.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx01.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx01.hwnd = NULL;
    		idx01.lpVerb = "runas";
    		idx01.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		//idx01.lpParameters = "/c winver";
    		idx01.lpParameters = "/c slmgr /ipk NW6C2-QMPVW-D7KKK-3GKT6-VCFB2";
    		idx01.lpDirectory = NULL;
    		idx01.nShow = SW_HIDE;
    		idx01.hInstApp = NULL;
    		ShellExecuteEx(&idx01);
    		WaitForSingleObject(idx01.hProcess, INFINITE);
			CloseHandle(idx01.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
			idx02.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx02.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx02.hwnd = NULL;
    		idx02.lpVerb = "runas";
    		idx02.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		//idx02.lpParameters = "/c winver";
    		idx02.lpParameters = "/c slmgr /skms kms8.msguides.com";
    		idx02.lpDirectory = NULL;
    		idx02.nShow = SW_HIDE;
    		idx02.hInstApp = NULL;
    		ShellExecuteEx(&idx02);
    		WaitForSingleObject(idx02.hProcess, INFINITE);
			CloseHandle(idx02.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
			idx03.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx03.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx03.hwnd = NULL;
    		idx03.lpVerb = "runas";
    		idx03.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		//idx03.lpParameters = "/c winver";
    		idx03.lpParameters = "/c slmgr /ato";
    		idx03.lpDirectory = NULL;
    		idx03.nShow = SW_HIDE;
    		idx03.hInstApp = NULL;
    		ShellExecuteEx(&idx03);
    		WaitForSingleObject(idx03.hProcess, INFINITE);
			CloseHandle(idx03.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
    		if(GetLastError() != 0){
    			FlashWindow(hwnd, TRUE);
    			MessageBox(hwnd, "An error occurred while activating Windows", "Activate Windows 10/11 x32", MB_ICONERROR|MB_APPLMODAL);
    			EnableWindow(hActiveB, TRUE);
    			EnableWindow(hActiveVerWin, TRUE);
			}
			else{
				FlashWindow(hwnd, TRUE);
    			msgboxid = MessageBox(hwnd, "Do you want to open activation settings?", "Activate Windows 10/11 x32", MB_ICONINFORMATION|MB_APPLMODAL|MB_YESNO);
			}
		}
		if((HWND) lParam == hActiveVerWin && index == 0){
   			SetWindowText(hActiveStringKey2, "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2");
   			SetWindowText(hActiveStringKMS2, "kms8.msguides.com");
   			EnableWindow(hActiveB, TRUE);
   		}
    	if((HWND) lParam == hActiveB && index == 1){
    		EnableWindow(hActiveB, FALSE);
    		EnableWindow(hActiveVerWin, FALSE);
    		idx11.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx11.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx11.hwnd = NULL;
    		idx11.lpVerb = "runas";
    		idx11.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx11.lpParameters = "/c slmgr /ipk 2WH4N-8QGBV-H22JP-CT43Q-MDWWJ";
    		idx11.lpDirectory = NULL;
    		idx11.nShow = SW_HIDE;
    		idx11.hInstApp = NULL;
    		ShellExecuteEx(&idx11);
    		WaitForSingleObject(idx11.hProcess, INFINITE);
			CloseHandle(idx11.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
			idx12.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx12.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx12.hwnd = NULL;
    		idx12.lpVerb = "runas";
    		idx12.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx12.lpParameters = "/c slmgr /skms kms8.msguides.com";
    		idx12.lpDirectory = NULL;
    		idx12.nShow = SW_HIDE;
    		idx12.hInstApp = NULL;
    		ShellExecuteEx(&idx12);
    		WaitForSingleObject(idx12.hProcess, INFINITE);
			CloseHandle(idx12.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
			idx13.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx13.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx13.hwnd = NULL;
    		idx13.lpVerb = "runas";
    		idx13.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx13.lpParameters = "/c slmgr /ato";
    		idx13.lpDirectory = NULL;
    		idx13.nShow = SW_HIDE;
    		idx13.hInstApp = NULL;
    		ShellExecuteEx(&idx13);
    		WaitForSingleObject(idx13.hProcess, INFINITE);
			CloseHandle(idx13.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
    		if(GetLastError() != 0){
    			FlashWindow(hwnd, TRUE);
    			MessageBox(hwnd, "An error occurred while activating Windows", "Activate Windows 10/11 x32", MB_ICONERROR|MB_APPLMODAL);
    			EnableWindow(hActiveB, TRUE);
    			EnableWindow(hActiveVerWin, TRUE);
			}
			else{
				FlashWindow(hwnd, TRUE);
    			msgboxid = MessageBox(hwnd, "Do you want to open activation settings?", "Activate Windows 10/11 x32", MB_ICONINFORMATION|MB_APPLMODAL|MB_YESNO);
			}
		}
		if((HWND) lParam == hActiveVerWin && index == 1){
   			SetWindowText(hActiveStringKey2, "2WH4N-8QGBV-H22JP-CT43Q-MDWWJ");
   			SetWindowText(hActiveStringKMS2, "kms8.msguides.com");
   			EnableWindow(hActiveB, TRUE);
   		}
    	if((HWND) lParam == hActiveB && index == 2){
    		EnableWindow(hActiveB, FALSE);
    		EnableWindow(hActiveVerWin, FALSE);
    		idx21.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx21.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx21.hwnd = NULL;
    		idx21.lpVerb = "runas";
    		idx21.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx21.lpParameters = "/c slmgr /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43";
    		idx21.lpDirectory = NULL;
    		idx21.nShow = SW_HIDE;
    		idx21.hInstApp = NULL;
    		ShellExecuteEx(&idx21);
    		WaitForSingleObject(idx21.hProcess, INFINITE);
			CloseHandle(idx21.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
			idx22.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx22.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx22.hwnd = NULL;
    		idx22.lpVerb = "runas";
    		idx22.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx22.lpParameters = "/c slmgr /skms kms8.msguides.com";
    		idx22.lpDirectory = NULL;
    		idx22.nShow = SW_HIDE;
    		idx22.hInstApp = NULL;
    		ShellExecuteEx(&idx22);
    		WaitForSingleObject(idx22.hProcess, INFINITE);
			CloseHandle(idx22.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
			idx23.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx23.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx23.hwnd = NULL;
    		idx23.lpVerb = "runas";
    		idx23.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx23.lpParameters = "/c slmgr /ato";
    		idx23.lpDirectory = NULL;
    		idx23.nShow = SW_HIDE;
    		idx23.hInstApp = NULL;
    		ShellExecuteEx(&idx23);
    		WaitForSingleObject(idx23.hProcess, INFINITE);
			CloseHandle(idx23.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
    		if(GetLastError() != 0){
    			FlashWindow(hwnd, TRUE);
    			MessageBox(hwnd, "An error occurred while activating Windows", "Activate Windows 10/11 x32", MB_ICONERROR|MB_APPLMODAL);
    			EnableWindow(hActiveB, TRUE);
    			EnableWindow(hActiveVerWin, TRUE);
			}
			else{
				FlashWindow(hwnd, TRUE);
    			msgboxid = MessageBox(hwnd, "Do you want to open activation settings?", "Activate Windows 10/11 x32", MB_ICONINFORMATION|MB_APPLMODAL|MB_YESNO);
			}
		}
		if((HWND) lParam == hActiveVerWin && index == 2){
   			SetWindowText(hActiveStringKey2, "NPPR9-FWDCX-D2C8J-H872K-2YT43");
   			SetWindowText(hActiveStringKMS2, "kms8.msguides.com");
   			EnableWindow(hActiveB, TRUE);
   		}
    	if((HWND) lParam == hActiveB && index == 3){
    		EnableWindow(hActiveB, FALSE);
    		EnableWindow(hActiveVerWin, FALSE);
    		idx31.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx31.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx31.hwnd = NULL;
    		idx31.lpVerb = "runas";
    		idx31.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx31.lpParameters = "/c slmgr /ipk DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4";
    		idx31.lpDirectory = NULL;
    		idx31.nShow = SW_HIDE;
    		idx31.hInstApp = NULL;
    		ShellExecuteEx(&idx31);
    		WaitForSingleObject(idx31.hProcess, INFINITE);
			CloseHandle(idx31.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
			idx32.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx32.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx32.hwnd = NULL;
    		idx32.lpVerb = "runas";
    		idx32.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx32.lpParameters = "/c slmgr /skms kms8.msguides.com";
    		idx32.lpDirectory = NULL;
    		idx32.nShow = SW_HIDE;
    		idx32.hInstApp = NULL;
    		ShellExecuteEx(&idx32);
    		WaitForSingleObject(idx32.hProcess, INFINITE);
			CloseHandle(idx32.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
			idx33.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx33.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx33.hwnd = NULL;
    		idx33.lpVerb = "runas";
    		idx33.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx33.lpParameters = "/c slmgr /ato";
    		idx33.lpDirectory = NULL;
    		idx33.nShow = SW_HIDE;
    		idx33.hInstApp = NULL;
    		ShellExecuteEx(&idx33);
    		WaitForSingleObject(idx33.hProcess, INFINITE);
			CloseHandle(idx33.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
    		if(GetLastError() != 0){
    			FlashWindow(hwnd, TRUE);
    			MessageBox(hwnd, "An error occurred while activating Windows", "Activate Windows 10/11 x32", MB_ICONERROR|MB_APPLMODAL);
    			EnableWindow(hActiveB, TRUE);
    			EnableWindow(hActiveVerWin, TRUE);
			}
			else{
				FlashWindow(hwnd, TRUE);
    			msgboxid = MessageBox(hwnd, "Do you want to open activation settings?", "Activate Windows 10/11 x32", MB_ICONINFORMATION|MB_APPLMODAL|MB_YESNO);
			}
		}
		if((HWND) lParam == hActiveVerWin && index == 3){
   			SetWindowText(hActiveStringKey2, "DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4");
   			SetWindowText(hActiveStringKMS2, "kms8.msguides.com");
   			EnableWindow(hActiveB, TRUE);
   		}
    	if((HWND) lParam == hActiveB && index == 4){
    		EnableWindow(hActiveB, FALSE);
    		EnableWindow(hActiveVerWin, FALSE);
    		idx41.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx41.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx41.hwnd = NULL;
    		idx41.lpVerb = "runas";
    		idx41.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx41.lpParameters = "/c slmgr /ipk PVMJN-6DFY6-9CCP6-7BKTT-D3WVR";
    		idx41.lpDirectory = NULL;
    		idx41.nShow = SW_HIDE;
    		idx41.hInstApp = NULL;
    		ShellExecuteEx(&idx41);
    		WaitForSingleObject(idx41.hProcess, INFINITE);
			CloseHandle(idx41.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
			idx42.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx42.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx42.hwnd = NULL;
    		idx42.lpVerb = "runas";
    		idx42.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx42.lpParameters = "/c slmgr /skms kms8.msguides.com";
    		idx42.lpDirectory = NULL;
    		idx42.nShow = SW_HIDE;
    		idx42.hInstApp = NULL;
    		ShellExecuteEx(&idx42);
    		WaitForSingleObject(idx42.hProcess, INFINITE);
			CloseHandle(idx42.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
			idx43.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx43.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx43.hwnd = NULL;
    		idx43.lpVerb = "runas";
    		idx43.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx43.lpParameters = "/c slmgr /ato";
    		idx43.lpDirectory = NULL;
    		idx43.nShow = SW_HIDE;
    		idx43.hInstApp = NULL;
    		ShellExecuteEx(&idx43);
    		WaitForSingleObject(idx43.hProcess, INFINITE);
			CloseHandle(idx43.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
    		if(GetLastError() != 0){
    			FlashWindow(hwnd, TRUE);
    			MessageBox(hwnd, "An error occurred while activating Windows", "Activate Windows 10/11 x32", MB_ICONERROR|MB_APPLMODAL);
    			EnableWindow(hActiveB, TRUE);
    			EnableWindow(hActiveVerWin, TRUE);
			}
			else{
				FlashWindow(hwnd, TRUE);
    			msgboxid = MessageBox(hwnd, "Do you want to open activation settings?", "Activate Windows 10/11 x32", MB_ICONINFORMATION|MB_APPLMODAL|MB_YESNO);
			}
		}
		if((HWND) lParam == hActiveVerWin && index == 4){
   			SetWindowText(hActiveStringKey2, "PVMJN-6DFY6-9CCP6-7BKTT-D3WVR");
   			SetWindowText(hActiveStringKMS2, "kms8.msguides.com");
   			EnableWindow(hActiveB, TRUE);
   		}
    	if((HWND) lParam == hActiveB && index == 5){
    		EnableWindow(hActiveB, FALSE);
    		EnableWindow(hActiveVerWin, FALSE);
    		idx51.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx51.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx51.hwnd = NULL;
    		idx51.lpVerb = "runas";
    		idx51.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx51.lpParameters = "/c slmgr /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99";
    		idx51.lpDirectory = NULL;
    		idx51.nShow = SW_HIDE;
    		idx51.hInstApp = NULL;
    		ShellExecuteEx(&idx51);
    		WaitForSingleObject(idx51.hProcess, INFINITE);
			CloseHandle(idx51.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
			idx52.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx52.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx52.hwnd = NULL;
    		idx52.lpVerb = "runas";
    		idx52.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx52.lpParameters = "/c slmgr /skms kms8.msguides.com";
    		idx52.lpDirectory = NULL;
    		idx52.nShow = SW_HIDE;
    		idx52.hInstApp = NULL;
    		ShellExecuteEx(&idx52);
    		WaitForSingleObject(idx52.hProcess, INFINITE);
			CloseHandle(idx52.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
			idx53.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx53.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx53.hwnd = NULL;
    		idx53.lpVerb = "runas";
    		idx53.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx53.lpParameters = "/c slmgr /ato";
    		idx53.lpDirectory = NULL;
    		idx53.nShow = SW_HIDE;
    		idx53.hInstApp = NULL;
    		ShellExecuteEx(&idx53);
    		WaitForSingleObject(idx53.hProcess, INFINITE);
			CloseHandle(idx53.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
    		if(GetLastError() != 0){
    			FlashWindow(hwnd, TRUE);
    			MessageBox(hwnd, "An error occurred while activating Windows", "Activate Windows 10/11 x32", MB_ICONERROR|MB_APPLMODAL);
    			EnableWindow(hActiveB, TRUE);
    			EnableWindow(hActiveVerWin, TRUE);
			}
			else{
				FlashWindow(hwnd, TRUE);
    			msgboxid = MessageBox(hwnd, "Do you want to open activation settings?", "Activate Windows 10/11 x32", MB_ICONINFORMATION|MB_APPLMODAL|MB_YESNO);
			}
		}
		if((HWND) lParam == hActiveVerWin && index == 5){
   			SetWindowText(hActiveStringKey2, "TX9XD-98N7V-6WMQ6-BX7FG-H8Q99");
   			SetWindowText(hActiveStringKMS2, "kms8.msguides.com");
   			EnableWindow(hActiveB, TRUE);
   		}
   		if((HWND) lParam == hActiveB && index == 6){
    		EnableWindow(hActiveB, FALSE);
    		EnableWindow(hActiveVerWin, FALSE);
    		idx61.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx61.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx61.hwnd = NULL;
    		idx61.lpVerb = "runas";
    		idx61.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx61.lpParameters = "/c slmgr /ipk 3KHY7-WNT83-DGQKR-F7HPR-844BM";
    		idx61.lpDirectory = NULL;
    		idx61.nShow = SW_HIDE;
    		idx61.hInstApp = NULL;
    		ShellExecuteEx(&idx61);
    		WaitForSingleObject(idx61.hProcess, INFINITE);
			CloseHandle(idx61.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
			idx62.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx62.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx62.hwnd = NULL;
    		idx62.lpVerb = "runas";
    		idx62.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx62.lpParameters = "/c slmgr /skms kms8.msguides.com";
    		idx62.lpDirectory = NULL;
    		idx62.nShow = SW_HIDE;
    		idx62.hInstApp = NULL;
    		ShellExecuteEx(&idx62);
    		WaitForSingleObject(idx62.hProcess, INFINITE);
			CloseHandle(idx62.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
			idx63.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx63.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx63.hwnd = NULL;
    		idx63.lpVerb = "runas";
    		idx63.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx63.lpParameters = "/c slmgr /ato";
    		idx63.lpDirectory = NULL;
    		idx63.nShow = SW_HIDE;
    		idx63.hInstApp = NULL;
    		ShellExecuteEx(&idx63);
    		WaitForSingleObject(idx63.hProcess, INFINITE);
			CloseHandle(idx63.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
    		if(GetLastError() != 0){
    			FlashWindow(hwnd, TRUE);
    			MessageBox(hwnd, "An error occurred while activating Windows", "Activate Windows 10/11 x32", MB_ICONERROR|MB_APPLMODAL);
    			EnableWindow(hActiveB, TRUE);
    			EnableWindow(hActiveVerWin, TRUE);
			}
			else{
				FlashWindow(hwnd, TRUE);
    			msgboxid = MessageBox(hwnd, "Do you want to open activation settings?", "Activate Windows 10/11 x32", MB_ICONINFORMATION|MB_APPLMODAL|MB_YESNO);
			}
		}
		if((HWND) lParam == hActiveVerWin && index == 6){
   			SetWindowText(hActiveStringKey2, "3KHY7-WNT83-DGQKR-F7HPR-844BM");
   			SetWindowText(hActiveStringKMS2, "kms8.msguides.com");
   			EnableWindow(hActiveB, TRUE);
   		}
   		if((HWND) lParam == hActiveB && index == 7){
    		EnableWindow(hActiveB, FALSE);
    		EnableWindow(hActiveVerWin, FALSE);
    		idx71.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx71.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx71.hwnd = NULL;
    		idx71.lpVerb = "runas";
    		idx71.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx71.lpParameters = "/c slmgr /ipk 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH";
    		idx71.lpDirectory = NULL;
    		idx71.nShow = SW_HIDE;
    		idx71.hInstApp = NULL;
    		ShellExecuteEx(&idx71);
    		WaitForSingleObject(idx71.hProcess, INFINITE);
			CloseHandle(idx71.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
			idx72.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx72.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx72.hwnd = NULL;
    		idx72.lpVerb = "runas";
    		idx72.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx72.lpParameters = "/c slmgr /skms kms8.msguides.com";
    		idx72.lpDirectory = NULL;
    		idx72.nShow = SW_HIDE;
    		idx72.hInstApp = NULL;
    		ShellExecuteEx(&idx72);
    		WaitForSingleObject(idx72.hProcess, INFINITE);
			CloseHandle(idx72.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
			idx73.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx73.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx73.hwnd = NULL;
    		idx73.lpVerb = "runas";
    		idx73.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx73.lpParameters = "/c slmgr /ato";
    		idx73.lpDirectory = NULL;
    		idx73.nShow = SW_HIDE;
    		idx73.hInstApp = NULL;
    		ShellExecuteEx(&idx73);
    		WaitForSingleObject(idx73.hProcess, INFINITE);
			CloseHandle(idx73.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
    		if(GetLastError() != 0){
    			FlashWindow(hwnd, TRUE);
    			MessageBox(hwnd, "An error occurred while activating Windows", "Activate Windows 10/11 x32", MB_ICONERROR|MB_APPLMODAL);
    			EnableWindow(hActiveB, TRUE);
    			EnableWindow(hActiveVerWin, TRUE);
			}
			else{
				FlashWindow(hwnd, TRUE);
    			msgboxid = MessageBox(hwnd, "Do you want to open activation settings?", "Activate Windows 10/11 x32", MB_ICONINFORMATION|MB_APPLMODAL|MB_YESNO);
			}
		}
		if((HWND) lParam == hActiveVerWin && index == 7){
   			SetWindowText(hActiveStringKey2, "7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH");
   			SetWindowText(hActiveStringKMS2, "kms8.msguides.com");
   			EnableWindow(hActiveB, TRUE);
   		}
   		if((HWND) lParam == hActiveB && index == 8){
    		EnableWindow(hActiveB, FALSE);
    		EnableWindow(hActiveVerWin, FALSE);
    		idx81.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx81.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx81.hwnd = NULL;
    		idx81.lpVerb = "runas";
    		idx81.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx81.lpParameters = "/c slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX";
    		idx81.lpDirectory = NULL;
    		idx81.nShow = SW_HIDE;
    		idx81.hInstApp = NULL;
    		ShellExecuteEx(&idx81);
    		WaitForSingleObject(idx81.hProcess, INFINITE);
			CloseHandle(idx81.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
			idx82.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx82.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx82.hwnd = NULL;
    		idx82.lpVerb = "runas";
    		idx82.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx82.lpParameters = "/c slmgr /skms kms8.msguides.com";
    		idx82.lpDirectory = NULL;
    		idx82.nShow = SW_HIDE;
    		idx82.hInstApp = NULL;
    		ShellExecuteEx(&idx82);
    		WaitForSingleObject(idx82.hProcess, INFINITE);
			CloseHandle(idx82.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
			idx83.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx83.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx83.hwnd = NULL;
    		idx83.lpVerb = "runas";
    		idx83.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx83.lpParameters = "/c slmgr /ato";
    		idx83.lpDirectory = NULL;
    		idx83.nShow = SW_HIDE;
    		idx83.hInstApp = NULL;
    		ShellExecuteEx(&idx83);
    		WaitForSingleObject(idx83.hProcess, INFINITE);
			CloseHandle(idx83.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
    		if(GetLastError() != 0){
    			FlashWindow(hwnd, TRUE);
    			MessageBox(hwnd, "An error occurred while activating Windows", "Activate Windows 10/11 x32", MB_ICONERROR|MB_APPLMODAL);
    			EnableWindow(hActiveB, TRUE);
    			EnableWindow(hActiveVerWin, TRUE);
			}
			else{
				FlashWindow(hwnd, TRUE);
    			msgboxid = MessageBox(hwnd, "Do you want to open activation settings?", "Activate Windows 10/11 x32", MB_ICONINFORMATION|MB_APPLMODAL|MB_YESNO);
			}
		}
		if((HWND) lParam == hActiveVerWin && index == 8){
   			SetWindowText(hActiveStringKey2, "W269N-WFGWX-YVC9B-4J6C9-T83GX");
   			SetWindowText(hActiveStringKMS2, "kms8.msguides.com");
   			EnableWindow(hActiveB, TRUE);
   		}
   		if((HWND) lParam == hActiveB && index == 9){
    		EnableWindow(hActiveB, FALSE);
    		EnableWindow(hActiveVerWin, FALSE);
    		idx91.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx91.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx91.hwnd = NULL;
    		idx91.lpVerb = "runas";
    		idx91.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx91.lpParameters = "/c slmgr /ipk MH37W-N47XK-V7XM9-C7227-GCQG9";
    		idx91.lpDirectory = NULL;
    		idx91.nShow = SW_HIDE;
    		idx91.hInstApp = NULL;
    		ShellExecuteEx(&idx91);
    		WaitForSingleObject(idx91.hProcess, INFINITE);
			CloseHandle(idx91.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
			idx92.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx92.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx92.hwnd = NULL;
    		idx92.lpVerb = "runas";
    		idx92.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx92.lpParameters = "/c slmgr /skms kms8.msguides.com";
    		idx92.lpDirectory = NULL;
    		idx92.nShow = SW_HIDE;
    		idx92.hInstApp = NULL;
    		ShellExecuteEx(&idx92);
    		WaitForSingleObject(idx92.hProcess, INFINITE);
			CloseHandle(idx92.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
			idx93.cbSize = sizeof(SHELLEXECUTEINFO);
    		idx93.fMask = SEE_MASK_NOCLOSEPROCESS;
    		idx93.hwnd = NULL;
    		idx93.lpVerb = "runas";
    		idx93.lpFile = "C:\\Windows\\system32\\cmd.exe";
    		idx93.lpParameters = "/c slmgr /ato";
    		idx93.lpDirectory = NULL;
    		idx93.nShow = SW_HIDE;
    		idx93.hInstApp = NULL;
    		ShellExecuteEx(&idx93);
    		WaitForSingleObject(idx93.hProcess, INFINITE);
			CloseHandle(idx93.hProcess);
			SendMessage(hActivePro, PBM_STEPIT, 0, 0);
			
    		if(GetLastError() != 0){
    			FlashWindow(hwnd, TRUE);
    			MessageBox(hwnd, "An error occurred while activating Windows", "Activate Windows 10/11 x32", MB_ICONERROR|MB_APPLMODAL);
    			EnableWindow(hActiveB, TRUE);
    			EnableWindow(hActiveVerWin, TRUE);
			}
			else{
				FlashWindow(hwnd, TRUE);
    			msgboxid = MessageBox(hwnd, "Do you want to open activation settings?", "Activate Windows 10/11 x32", MB_ICONINFORMATION|MB_APPLMODAL|MB_YESNO);
			}
		}
		if((HWND) lParam == hActiveVerWin && index == 9){
   			SetWindowText(hActiveStringKey2, "MH37W-N47XK-V7XM9-C7227-GCQG9");
   			SetWindowText(hActiveStringKMS2, "kms8.msguides.com");
   			EnableWindow(hActiveB, TRUE);
   		}
   		if((HWND) lParam == hActiveBCheck){
   			ShellExecute(NULL, "open", "C:\\Windows\\system32\\cmd.exe", "/c slmgr /dli && slmgr /dlv && slmgr /xpr", NULL, SW_HIDE);
		   }
		if((HWND) lParam == hActiveAbout){
			MessageBox(hwnd, " Activate Windows 10/11 x32 by bako35 \n Ver. 1.0.0.0 \n github.com/bako35", "About", MB_ICONINFORMATION|MB_APPLMODAL);
		}
		   
		switch (msgboxid)
		{
			case IDYES:
				ShellExecute(NULL, "open", "ms-settings:activation", NULL, NULL, SW_NORMAL);
				DestroyWindow(hwnd);
				break;
			case IDNO:
				DestroyWindow(hwnd);
				break;
		}
        default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
        break;
    }
    return 0;
}


