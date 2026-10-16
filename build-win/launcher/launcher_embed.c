// Timesheets launcher (1.0.1): runs the app INSIDE this process by loading the
// bundled Python DLL, so Windows only ever sees Timesheets.exe (our icon, our
// name in Task Manager, pins correctly). If the DLL can't be loaded it falls
// back to the 1.0.0 behaviour of starting runtime\pythonw.exe.
#include <windows.h>
#include <shlwapi.h>

typedef int (*PyMainFn)(int, wchar_t **);

static void join(wchar_t *out, const wchar_t *a, const wchar_t *b) {
    lstrcpyW(out, a); lstrcatW(out, L"\\"); lstrcatW(out, b);
}

static int fallback(const wchar_t *dir, const wchar_t *py, const wchar_t *script) {
    static wchar_t cmdline[3 * MAX_PATH];
    lstrcpyW(cmdline, L"\""); lstrcatW(cmdline, py); lstrcatW(cmdline, L"\" \"");
    lstrcatW(cmdline, script); lstrcatW(cmdline, L"\"");
    STARTUPINFOW si; PROCESS_INFORMATION pi;
    ZeroMemory(&si, sizeof(si)); si.cb = sizeof(si); ZeroMemory(&pi, sizeof(pi));
    if (!CreateProcessW(py, cmdline, NULL, NULL, FALSE, CREATE_NO_WINDOW, NULL, dir, &si, &pi)) {
        MessageBoxW(NULL, L"Timesheets could not start. Please reinstall Timesheets.", L"Timesheets", MB_OK | MB_ICONERROR);
        return 1;
    }
    CloseHandle(pi.hThread); CloseHandle(pi.hProcess);
    return 0;
}

int WINAPI wWinMain(HINSTANCE hInst, HINSTANCE hPrev, PWSTR cmd, int show) {
    static wchar_t dir[MAX_PATH], rt[MAX_PATH], py[MAX_PATH], dll[MAX_PATH], script[MAX_PATH], self[MAX_PATH];
    GetModuleFileNameW(NULL, self, MAX_PATH);
    lstrcpyW(dir, self); PathRemoveFileSpecW(dir);
    join(rt, dir, L"runtime");
    join(py, rt, L"pythonw.exe");
    join(dll, rt, L"python312.dll");
    join(script, dir, L"app\\timesheets_client.py");
    if (GetFileAttributesW(dll) == INVALID_FILE_ATTRIBUTES) {
        wchar_t msg[2 * MAX_PATH];
        lstrcpyW(msg, L"The Timesheets runtime is missing. Please reinstall Timesheets.\n\nLooked for:\n");
        lstrcatW(msg, dll);
        MessageBoxW(NULL, msg, L"Timesheets", MB_OK | MB_ICONERROR);
        return 1;
    }
    SetCurrentDirectoryW(dir);
    SetEnvironmentVariableW(L"PYTHONHOME", rt);
    SetEnvironmentVariableW(L"PYTHONPATH", NULL);
    SetEnvironmentVariableW(L"PYTHONNOUSERSITE", L"1");
    SetEnvironmentVariableW(L"TIMESHEETS_EMBEDDED", L"1");
    SetDllDirectoryW(rt);
    HMODULE h = LoadLibraryExW(dll, NULL, LOAD_WITH_ALTERED_SEARCH_PATH);
    PyMainFn pymain = h ? (PyMainFn)GetProcAddress(h, "Py_Main") : NULL;
    if (!pymain) {
        SetEnvironmentVariableW(L"PYTHONHOME", NULL);
        SetEnvironmentVariableW(L"TIMESHEETS_EMBEDDED", NULL);
        return fallback(dir, py, script);
    }
    wchar_t *argv[] = { self, script, NULL };
    return pymain(2, argv);
}
