// Timesheets launcher: runs <exe dir>\runtime\pythonw.exe <exe dir>\app\timesheet_app.py
// with no console window. Any failure is shown in a message box.
#include <windows.h>
#include <shlwapi.h>

static void join(wchar_t *out, const wchar_t *a, const wchar_t *b) {
    lstrcpyW(out, a); lstrcatW(out, L"\\"); lstrcatW(out, b);
}

int WINAPI wWinMain(HINSTANCE hInst, HINSTANCE hPrev, PWSTR cmd, int show) {
    static wchar_t dir[MAX_PATH], py[MAX_PATH], script[MAX_PATH], cmdline[3 * MAX_PATH];
    GetModuleFileNameW(NULL, dir, MAX_PATH);
    PathRemoveFileSpecW(dir);
    join(py, dir, L"runtime\\pythonw.exe");
    join(script, dir, L"app\\timesheet_app.py");
    if (GetFileAttributesW(py) == INVALID_FILE_ATTRIBUTES) {
        wchar_t msg[2 * MAX_PATH];
        lstrcpyW(msg, L"The Timesheets runtime is missing. Please reinstall Timesheets.\n\nLooked for:\n");
        lstrcatW(msg, py);
        MessageBoxW(NULL, msg, L"Timesheets", MB_OK | MB_ICONERROR);
        return 1;
    }
    lstrcpyW(cmdline, L"\""); lstrcatW(cmdline, py); lstrcatW(cmdline, L"\" \"");
    lstrcatW(cmdline, script); lstrcatW(cmdline, L"\"");
    STARTUPINFOW si; PROCESS_INFORMATION pi;
    ZeroMemory(&si, sizeof(si)); si.cb = sizeof(si);
    ZeroMemory(&pi, sizeof(pi));
    if (!CreateProcessW(py, cmdline, NULL, NULL, FALSE, CREATE_NO_WINDOW, NULL, dir, &si, &pi)) {
        MessageBoxW(NULL, L"Timesheets could not start. Please reinstall Timesheets.", L"Timesheets", MB_OK | MB_ICONERROR);
        return 1;
    }
    CloseHandle(pi.hThread); CloseHandle(pi.hProcess);
    return 0;
}
