// to use rand_s()
#define _CRT_RAND_S

//#define _WIN32_WINNT _WIN32_WINNT_WINXP 
//#define WDK_NTDDI_VERSION NTDDI_WINXPSP1

#include <opencv2/imgproc.hpp>
#include <opencv2/imgcodecs.hpp>
//#include <opencv2/highgui.hpp>

#include <Windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <io.h>
#include <VersionHelpers.h>
#include <algorithm>
#include <vector>
#include <math.h>

static constexpr const auto BUSYSTATE_CPU_USAGE_THRESHOLD  = 40;
static constexpr const auto BUSYSTATE_CPU_USAGE_CHK_INTVAL = 5;

static constexpr const auto BUSYSTATE_NOT_BUSY             = 0;
static constexpr const auto BUSYSTATE_BUSY                 = 1;

static constexpr const auto PIC_LIST_CHANGED               = 0;
static constexpr const auto PIC_LIST_NOT_CHANGED           = 1;

static constexpr const auto PROC_WALLPAPER_DONE            = 0;
static constexpr const auto PROC_WALLPAPER_DUP_PICTURE     = 1;

static constexpr const auto MIN_EMPTY_SPACE_COLOR          = 0;
static constexpr const auto EMPTY_SPACE_COLOR_B            = 0;
static constexpr const auto EMPTY_SPACE_COLOR_W            = 1;
static constexpr const auto EMPTY_SPACE_COLOR_D            = 2;
static constexpr const auto MAX_EMPTY_SPACE_COLOR          = 2;

static constexpr const auto MIN_PERIOD_IN_MINUTE           = 15;
static constexpr const auto MAX_PERIOD_IN_MINUTE           = 1440;

static constexpr const auto MIN_UPSCALE_MODE               = 0;
static constexpr const auto UPSCALE_MODE_DONT_UPSCALE      = 0;
static constexpr const auto UPSCALE_MODE_UPSCALE_UP_TO_2X  = 1;
static constexpr const auto UPSCALE_MODE_UPSCALE_UP_TO_4X  = 2;
static constexpr const auto UPSCALE_MODE_UPSCALE_SCR       = 3;
static constexpr const auto MAX_UPSCALE_MODE               = 3;

static constexpr const auto CONF_DEFAULT_EMPTY_SPACE_COLOR = EMPTY_SPACE_COLOR_D;
static constexpr const auto CONF_DEFAULT_PERIOD_IN_MINUTE  = 30;
static constexpr const auto CONF_DEFAULT_UPSCALE_MODE      = UPSCALE_MODE_UPSCALE_UP_TO_2X;

static constexpr const auto MAX_CONFIG_FILE_LENGTH         = 1536;
static constexpr const auto MAX_PICTURE_NUMBER             = 100000;
static constexpr const auto MAX_IMAGE_EXT_LENGTH           = 4;

static constexpr const BYTE      BOM_UTF8[]                 = { 0xEF, 0xBB, 0xBF };

static constexpr const ULONGLONG MAX_PICTURE_FILESIZE       = 1024ULL * 1024ULL * 128ULL;
static constexpr const char*     STR_MAX_PICTURE_FILESIZE   = "128MB";

static constexpr const wchar_t*  WSTR_PROGRAM_NAME          = L"FitWallpaper.exe";
static constexpr const char*     STR_WALLPAPER_FNAME_PNG    = "wallpaper.png";
static constexpr const wchar_t*  WSTR_WALLPAPER_FNAME_PNG   = L"wallpaper.png";
static constexpr const char*     STR_WALLPAPER_FNAME_JPEG   = "wallpaper.jpg";
static constexpr const wchar_t*  WSTR_WALLPAPER_FNAME_JPEG  = L"wallpaper.jpg";

static constexpr const wchar_t*  WSTR_CMDLINE_UNREGISTER    = L"/u";
static constexpr const wchar_t*  WSTR_CMDLINE_BOOT          = L"/boot";

static constexpr const char*     STR_CONF_DIR_PICTURE       = "dirPicture";
static constexpr const auto      LEN_CONF_DIR_PICTURE       = 10;
static constexpr const char*     STR_CONF_EMPTY_SPACE_COLOR = "emptySpaceColor";
static constexpr const auto      LEN_CONF_EMPTY_SPACE_COLOR = 15;
static constexpr const char*     STR_CONF_PERIOD_IN_MINUTE  = "periodInMinute";
static constexpr const auto      LEN_CONF_PERIOD_IN_MINUTE  = 14;
static constexpr const char*     STR_CONF_UPSCALE_MODE      = "upscaleMode";
static constexpr const auto      LEN_CONF_UPSCALE_MODE      = 11;

static constexpr const char*     STR_IMAGE_EXT_LIST[]       = {  "jpeg",  "jpg",  "jpe",  "png",  "bmp",  "tif",  "tiff" };
static constexpr const wchar_t*  WSTR_IMAGE_EXT_LIST[]      = { L"jpeg", L"jpg", L"jpe", L"png", L"bmp", L"tif", L"tiff" };

using namespace cv;
using namespace std;

int       isSystemBusy();
int       updatePictureList(const wchar_t* wDirPicture, wchar_t* picList, int& sizePicList);
int       processWallpaper(const wchar_t* picList, const int sizePicList, const bool bPicListChanged, const int emptySpaceColor, const int upscaleMode, const bool bUseJPEGFormat);

void      DisplayInfoBoxA(LPCSTR szInfoMsg);
void      DisplayErrorBoxW(LPCWSTR wszErrorMsg);

int       registerRunAtStartupReg(const wchar_t* wszDirWorking);
int       unregisterRunAtStartupReg();
int       updateJPEGImportQualityReg();
ULONGLONG getLastChangedReg();
int       updateLastChangedReg();
int       unregisterKeyFromReg();

int       checkDirectory(const wchar_t* dir, const wchar_t* dirName);
int       checkFile(const wchar_t* file, const wchar_t* fileName);
int       isSupportedImageFile(const wchar_t* imageFileName);

int       cvt_utf8_to_wchar(const char* src, const int srcSize, wchar_t* dest, const int destSize);
int       cvt_wchar_to_utf8(const wchar_t* src, const int srcSize, char* dest, const int destSize);
int       isIllegalCharacterExist(const wchar_t* wstr, const int size);

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    // require at least Windows XP SP1
    if (!IsWindowsXPSP1OrGreater()) {
        MessageBoxW(NULL, L"Windows version is too old!\nRequire at least XP SP1!", L"FitWallpaper: Error", MB_ICONERROR | MB_OK);
        return -1;
    }

    // if program run with WSTR_CMDLINE_UNREGISTER(/u)
    // try to unregister all registry and exit.
    if (wcslen(WSTR_CMDLINE_UNREGISTER) == wcslen(lpCmdLine) &&
        0 == wcsncmp(lpCmdLine, WSTR_CMDLINE_UNREGISTER, wcslen(WSTR_CMDLINE_UNREGISTER)))
    {
        if (-1 == unregisterRunAtStartupReg()) return -1;
        if (-1 == unregisterKeyFromReg()) return -1;
        return 0;
    }

    // check FitWallpaper::Running mutex for single instance
    const HANDLE hMutexRunning = CreateMutexA(NULL, TRUE, "FitWallpaper::Running");
    if (NULL == hMutexRunning) {
        DisplayErrorBoxW(L"Failed to create mutex!");
        return -1;
    }
    bool bRunning = false;
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
        bRunning = true;
        CloseHandle(hMutexRunning);
    }

    wchar_t wDirWorking[MAX_PATH] = L"";
    wchar_t wDirPicture[MAX_PATH] = L"";
    wchar_t wWallpaperPath[MAX_PATH] = L"";
    int emptySpaceColor = CONF_DEFAULT_EMPTY_SPACE_COLOR;
    int periodInMinute = CONF_DEFAULT_PERIOD_IN_MINUTE;
    int upscaleMode = CONF_DEFAULT_UPSCALE_MODE;
    bool bUseJPEGFormat = true;
    
    // Windows 8 or later use PNG format which is natively supported
    if (IsWindows8OrGreater()) {
        bUseJPEGFormat = false;
    }

    // check current directory
    {
        // get program (exe file) location including program file name.
        const DWORD lenDirWorking = GetModuleFileNameW(NULL, wDirWorking, MAX_PATH);

        // check error or overflow
        if (0 == lenDirWorking || MAX_PATH == lenDirWorking) {
            DisplayErrorBoxW(L"Failed to get a program location or a program location is too long!");
            return -1;
        }

        // delete characters from end to lastest '\'
        // which written as '\\FitWallpaper.exe'
        wchar_t* pEnd = wcsrchr(wDirWorking, L'\\');
        if (pEnd == NULL) {
            DisplayErrorBoxW(L"A current directory is invalid!");
            return -1;
        }

        // check program filename is correct
        wchar_t wszInvalidProgramName[128] = L"";
        if (-1 == swprintf_s(wszInvalidProgramName,
            L"A program filename is incorrect!\n"
            L"Correct one is <%s>!",
            WSTR_PROGRAM_NAME))
        {
            DisplayErrorBoxW(L"Failed to create error message say invalid program filename!");
            return -1;
        }

        const size_t lenProgramName = wcslen(WSTR_PROGRAM_NAME);
        if (lenProgramName != wcsnlen_s(pEnd + 1, lenProgramName + 1) ||
            0 != wcsncmp(pEnd + 1, WSTR_PROGRAM_NAME, lenProgramName)) {
            DisplayErrorBoxW(wszInvalidProgramName);
            return -1;
        }

        // fill '\\FitWallpaper.exe' as 0
        while (*pEnd != L'\0') {
            *pEnd = L'\0';
            pEnd++;
        }

        // check length
        if (wcslen(wDirWorking) >= (MAX_PATH - wcslen(WSTR_WALLPAPER_FNAME_JPEG) - 1 - 1)) {
            DisplayErrorBoxW(L"A current directory is too long!");
            return -1;
        }

        if (-1 == checkDirectory(wDirWorking, L"current directory")) return -1;

        // change current directory to program (exe file) location
        if (0 == SetCurrentDirectoryW(wDirWorking)) {
            DisplayErrorBoxW(L"Failed to change current directory!");
            return -1;
        }

        wcscat_s(wWallpaperPath, wDirWorking);
        wcscat_s(wWallpaperPath, L"\\");
        if (bUseJPEGFormat)
            wcscat_s(wWallpaperPath, WSTR_WALLPAPER_FNAME_JPEG);
        else
            wcscat_s(wWallpaperPath, WSTR_WALLPAPER_FNAME_PNG);
    }

    // update run at startup registry
    // only if lpCmdLine is empty and there are no instance running
    if (0 == wcslen(lpCmdLine) && !bRunning) {
        if (-1 == registerRunAtStartupReg(wDirWorking)) return -1;
        DisplayInfoBoxA("Run at startup registry is updated.\nPlease, run a program again if program's location is changed.");

        // multi-monitor check
        if (1 < GetSystemMetrics(SM_CMONITORS))
            DisplayInfoBoxA("This program doesn't support multi-monitor\nYou may continue to use this program\nBut how wallpaper be applied is unpredictable");

        // JPEGImportQuality registry check
        if (-1 == updateJPEGImportQualityReg()) return -1;

        // Theme roaming check
        if (IsWindows8OrGreater())
            DisplayInfoBoxA(
                "You may turn off <Theme roaming> in [Settings App -> User Accounts -> Sync your settings]\n"
                "Because there are compression for roamed themes if it exceed a limit");
    }

    // file operation
    {
        FILE* fp = NULL;
        // create file to stop program (stop.cmd)
        fopen_s(&fp, "stop.cmd", "rb");
        // if failed to open stop.cmd, make it because it doesn't exist.
        if (fp == NULL) {
            int err = GetLastError();
            if (err == ERROR_FILE_NOT_FOUND) {
                fopen_s(&fp, "stop.cmd", "wb");
                if (fp == NULL) {
                    DisplayErrorBoxW(L"Failed to create stop.cmd!");
                    return -1;
                }

                const char* szStopCmd =
                    "@echo off\n"
                    "echo kill program running\n"
                    "taskkill /f /im FitWallpaper.exe\n"
                    "echo remove startup registry\n"
                    "FitWallpaper.exe /u\n"
                    "echo press any key to exit\n"
                    "pause>nul";
                if (-1 == fwrite(szStopCmd, sizeof(char), strlen(szStopCmd), fp)) {
                    fclose(fp);
                    DisplayErrorBoxW(L"Failed to write stop.cmd!");
                    return -1;
                }
            }
            else {
                DisplayErrorBoxW(L"Failed to open stop.cmd!");
                return -1;
            }
        }
        fclose(fp);

        // load config file (config.txt)
        fopen_s(&fp, "config.txt", "rb");
        // if failed to open config.txt, make default config file because it doesn't exist. and exit.
        if (fp == NULL) {
            int err = GetLastError();
            if (err == ERROR_FILE_NOT_FOUND) {
                fopen_s(&fp, "config.txt", "wb");
                if (fp == NULL) {
                    DisplayErrorBoxW(L"Failed to create config.txt!");
                    return -1;
                }

                if (-1 == fwrite(BOM_UTF8, sizeof(BYTE), 3, fp)) {
                    fclose(fp);
                    DisplayErrorBoxW(L"Failed to write config.txt!");
                    return -1;
                }

                char szImageExtList[MAX_CONFIG_FILE_LENGTH] = "";

                for (const char* szImageExt : STR_IMAGE_EXT_LIST) {
                    strcat_s(szImageExtList, szImageExt);
                    strcat_s(szImageExtList, ", ");
                }
                szImageExtList[strnlen_s(szImageExtList, MAX_CONFIG_FILE_LENGTH) - 1] = '\0';
                szImageExtList[strnlen_s(szImageExtList, MAX_CONFIG_FILE_LENGTH) - 1] = '\0';

                char szConfDefault[MAX_CONFIG_FILE_LENGTH] = "";
                if (-1 == sprintf_s(szConfDefault,
                    "# Please save this as UTF8 encoding\n"
                    "# Picture Directory - up to %d Pictures / each file up to %s\n"
                    "# Support format: %s\n"
                    "%s = D:\\Pictures\n"
                    "\n"
                    "# Fill an empty space with selected color (Default: %d)\n"
                    "# 0: Black, 1: White, 2: Dominant Color\n"
                    "%s = %d\n"
                    "\n"
                    "# Change picture every X minute(s) [15 ~ 1440] (Default: %d)\n"
                    "%s = %d\n"
                    "\n"
                    "# Upscale mode (Default: %d)\n"
                    "# 0: Don't upscale\n"
                    "# 1: Upscale picture up to 2x\n"
                    "# 2: Upscale picture up to 4x\n"
                    "# 3: Upscale picture to the screen size\n"
                    "%s = %d",
                    MAX_PICTURE_NUMBER, STR_MAX_PICTURE_FILESIZE,
                    szImageExtList,
                    STR_CONF_DIR_PICTURE,
                    CONF_DEFAULT_EMPTY_SPACE_COLOR,
                    STR_CONF_EMPTY_SPACE_COLOR, CONF_DEFAULT_EMPTY_SPACE_COLOR,
                    CONF_DEFAULT_PERIOD_IN_MINUTE,
                    STR_CONF_PERIOD_IN_MINUTE, CONF_DEFAULT_PERIOD_IN_MINUTE,
                    CONF_DEFAULT_UPSCALE_MODE,
                    STR_CONF_UPSCALE_MODE, CONF_DEFAULT_UPSCALE_MODE))
                {
                    fclose(fp);
                    DisplayErrorBoxW(L"Failed to create default config.txt data!");
                    return -1;
                }

                if (-1 == fwrite(szConfDefault, sizeof(char), strlen(szConfDefault), fp)) {
                    fclose(fp);
                    DisplayErrorBoxW(L"Failed to write config.txt!");
                    return -1;
                }

                fclose(fp);
                DisplayInfoBoxA("A config.txt is created in current directory!\nPlease, edit it properly and restart a program!");
                // exit without error
                return 0;
            }
            else {
                DisplayErrorBoxW(L"Failed to open config.txt!");
                return -1;
            }
        }

        // read config
        char pBuf[MAX_CONFIG_FILE_LENGTH] = "";
        size_t rdBytes = fread(pBuf, sizeof(BYTE), MAX_CONFIG_FILE_LENGTH, fp);
        fclose(fp);

        if (rdBytes >= MAX_CONFIG_FILE_LENGTH) {
            DisplayErrorBoxW(L"A config.txt file's size is too big!");
            return -1;
        }
        else if (rdBytes < LEN_CONF_DIR_PICTURE) {
            DisplayErrorBoxW(L"A config.txt file's size is too small!");
            return -1;
        }

        char* pStart;
        char* pEnd;
        char* pCur;
        char dirPicture[MAX_PATH] = "";
        pCur = &pBuf[0];

        if (0 == strncmp((const char*)BOM_UTF8, pCur, 3))
            pCur += 3;

        pCur--;
        do {
            // read line
            pStart = pCur + 1;
            pCur = strchr(pStart, '\r');
            if (pCur == NULL) {
                pCur = strchr(pStart, '\n');
                if (pCur == NULL) {
                    pCur = strchr(pStart, '\0');
                    if (pCur == NULL) {
                        DisplayErrorBoxW(L"A config.txt is damaged!");
                        return -1;
                    }
                }
            }
            pEnd = pCur;
            if (*pCur == '\r' && *(pCur + 1) == '\n') pCur++;

            // trim space and tab
            while (*pStart == ' ' || *pStart == '\t') pStart++;
            pEnd--;
            while (*pEnd == ' ' || *pEnd == '\t') pEnd--;
            pEnd++;

            // this line is not a comment
            if (*pStart != '#' && *pStart != '\0') {
                *pEnd = '\0';
                const size_t lenCurConf = strlen(pStart);
                // dirPicture
                if (lenCurConf > LEN_CONF_DIR_PICTURE &&
                    0 == strncmp(pStart, STR_CONF_DIR_PICTURE, LEN_CONF_DIR_PICTURE))
                {
                    pStart += LEN_CONF_DIR_PICTURE;
                    while (*pStart == ' ' || *pStart == '\t') pStart++;
                    if (*pStart != '=') continue;
                    pStart++;
                    while (*pStart == ' ' || *pStart == '\t') pStart++;

                    const size_t lenDirPicture = strnlen_s(pStart, MAX_PATH);
                    if (lenDirPicture == 0) {
                        DisplayErrorBoxW(L"A dirPicture value is empty!");
                        return -1;
                    }
                    else if (lenDirPicture >= MAX_PATH) {
                        DisplayErrorBoxW(L"A dirPicture value's length is too long!");
                        return -1;
                    }
                    else if (0 != memcpy_s(dirPicture, MAX_PATH, pStart, lenDirPicture)) {
                        DisplayErrorBoxW(L"Failed to get dirPicture value!");
                        return -1;
                    }

                    if (-1 == cvt_utf8_to_wchar(dirPicture, MAX_PATH, wDirPicture, MAX_PATH)) return -1;
                }
                // emptySpaceColor
                else if (lenCurConf > LEN_CONF_EMPTY_SPACE_COLOR &&
                    0 == strncmp(pStart, STR_CONF_EMPTY_SPACE_COLOR, LEN_CONF_EMPTY_SPACE_COLOR))
                {
                    pStart += LEN_CONF_EMPTY_SPACE_COLOR;
                    while (*pStart == ' ' || *pStart == '\t') pStart++;
                    if (*pStart != '=') continue;
                    pStart++;
                    while (*pStart == ' ' || *pStart == '\t') pStart++;

                    const size_t lenEsColor = strnlen_s(pStart, 2);
                    if (lenEsColor == 0) {
                        DisplayErrorBoxW(L"A emptySpaceColor value is empty!");
                        return -1;
                    }
                    else if (lenEsColor == 2) {
                        DisplayErrorBoxW(L"A emptySpaceColor value's length must be 1!");
                        return -1;
                    }

                    emptySpaceColor = *pStart - '0';
                    if (emptySpaceColor < MIN_EMPTY_SPACE_COLOR || emptySpaceColor > MAX_EMPTY_SPACE_COLOR) {
                        DisplayErrorBoxW(L"A emptySpaceColor value must be between 0 and 2 inclusive!");
                        return -1;
                    }
                }
                // periodInMinute
                else if (lenCurConf > LEN_CONF_PERIOD_IN_MINUTE &&
                    0 == strncmp(pStart, STR_CONF_PERIOD_IN_MINUTE, LEN_CONF_PERIOD_IN_MINUTE))
                {
                    pStart += LEN_CONF_PERIOD_IN_MINUTE;
                    while (*pStart == ' ' || *pStart == '\t') pStart++;
                    if (*pStart != '=') continue;
                    pStart++;
                    while (*pStart == ' ' || *pStart == '\t') pStart++;

                    char szMinute[5] = "";
                    const size_t lenPeriodInMinute = strnlen_s(pStart, 5);
                    if (lenPeriodInMinute == 0) {
                        DisplayErrorBoxW(L"A periodInMinute value is empty!");
                        return -1;
                    }
                    else if (lenPeriodInMinute == 1 || lenPeriodInMinute == 5) {
                        DisplayErrorBoxW(L"A periodInMinute value's length must be between 2 and 4 inclusive!");
                        return -1;
                    }

                    if (0 != memcpy_s(szMinute, 5, pStart, lenPeriodInMinute)) {
                        DisplayErrorBoxW(L"Failed to get periodInMinute value!");
                        return -1;
                    }

                    periodInMinute = atoi(szMinute);
                    if (periodInMinute == 0) {
                        DisplayErrorBoxW(L"A periodInMinute value is zero or not a number!");
                        return -1;
                    }
                    else if (periodInMinute < MIN_PERIOD_IN_MINUTE || periodInMinute > MAX_PERIOD_IN_MINUTE) {
                        DisplayErrorBoxW(L"A periodInMinute value must be between 15 and 1440 inclusive!");
                        return -1;
                    }
                }
                // upscaleMode
                else if (lenCurConf > LEN_CONF_UPSCALE_MODE &&
                    0 == strncmp(pStart, STR_CONF_UPSCALE_MODE, LEN_CONF_UPSCALE_MODE))
                {
                    pStart += LEN_CONF_UPSCALE_MODE;
                    while (*pStart == ' ' || *pStart == '\t') pStart++;
                    if (*pStart != '=') continue;
                    pStart++;
                    while (*pStart == ' ' || *pStart == '\t') pStart++;

                    const size_t lenUpscaleMode = strnlen_s(pStart, 2);
                    if (lenUpscaleMode == 0) {
                        DisplayErrorBoxW(L"A upscaleMode value is empty!");
                        return -1;
                    }
                    else if (lenUpscaleMode == 2) {
                        DisplayErrorBoxW(L"A upscaleMode value's length must be 1!");
                        return -1;
                    }

                    upscaleMode = *pStart - '0';
                    if (upscaleMode < MIN_UPSCALE_MODE || upscaleMode > MAX_UPSCALE_MODE) {
                        DisplayErrorBoxW(L"A upscaleMode value must be between 0 and 3 inclusive!");
                        return -1;
                    }
                }

                *pEnd = '\n';
            }
        } while (*pCur != '\0');
    }

    const bool bBootCmd = (wcslen(WSTR_CMDLINE_BOOT) == wcslen(lpCmdLine) &&
        0 == wcsncmp(lpCmdLine, WSTR_CMDLINE_BOOT, wcslen(WSTR_CMDLINE_BOOT)));

    // if lpCmdLine is file, try to update wallpaper using it and exit
    if (0 != wcslen(lpCmdLine) && !bBootCmd) {
        // trim cmdLine first
        // if filename has space, " is attached on front and end side
        // so remove them, too.
        wchar_t* pStart = lpCmdLine;
        wchar_t* pEnd = pStart + wcslen(lpCmdLine);

        while (*pStart == ' ' || *pStart == '\t') pStart++;
        if (L'\"' == *pStart) pStart++;

        pEnd--;
        while (*pEnd == ' ' || *pEnd == '\t') pEnd--;
        if (L'\"' == *pEnd) pEnd--;
        pEnd++;

        *pEnd = L'\0';

        // check file
        if (MAX_PATH <= wcsnlen_s(pStart, MAX_PATH)) {
            DisplayErrorBoxW(L"File in command line is too long or including multiple files!");
            return -1;
        }
        if (-1 == checkFile(pStart, L"picture file in argument")) return -1;
        if (-1 == isSupportedImageFile(pStart)) return -1;

        // check FitWallpaper::Processing mutex existance
        // to exit when image is processing in another instance
        HANDLE hMutexProcessing = CreateMutexA(NULL, TRUE, "FitWallpaper::Processing");
        if (NULL == hMutexProcessing) {
            DisplayErrorBoxW(L"Failed to create mutex!");
            return -1;
        }
        if (GetLastError() == ERROR_ALREADY_EXISTS) {
            CloseHandle(hMutexProcessing);
            DisplayInfoBoxA("Please retry after a wallpaper is changed! It's soon!");
            return 0;
        }
        CloseHandle(hMutexProcessing);

        if (-1 == processWallpaper(pStart, 1, true, emptySpaceColor, upscaleMode, bUseJPEGFormat)) return -1;

        // set wallpaper and exit
        SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, (void*)wWallpaperPath, SPIF_UPDATEINIFILE);
        return 0;
    }

    // check working directory
    if (-1 == checkDirectory(wDirPicture, L"picture directory")) return -1;

    // allocate picList in heap
    HANDLE hDefaultHeap = GetProcessHeap();
    wchar_t* picList = nullptr;
    int sizePicList = 0;

    if (nullptr == hDefaultHeap) {
        DisplayErrorBoxW(L"Failed to get default heap handle!");
        return -1;
    }
    picList = (wchar_t*)HeapAlloc(hDefaultHeap, HEAP_ZERO_MEMORY, MAX_PICTURE_NUMBER * MAX_PATH * sizeof(wchar_t));
    if (nullptr == picList) {
        DisplayErrorBoxW(L"Failed to allocate memory for picture list!");
        return -1;
    }

    // if this instance started while another instance is running
    // update wallpaper once and exit
    if (bRunning) {
        // check FitWallpaper::Processing mutex existance
        // to exit when image is processing in another instance
        HANDLE hMutexProcessing = CreateMutexA(NULL, TRUE, "FitWallpaper::Processing");
        if (NULL == hMutexProcessing) {
            DisplayErrorBoxW(L"Failed to create mutex!");
            HeapFree(GetProcessHeap(), 0, picList);
            return -1;
        }
        if (GetLastError() == ERROR_ALREADY_EXISTS) {
            CloseHandle(hMutexProcessing);
            HeapFree(GetProcessHeap(), 0, picList);
            DisplayInfoBoxA("Please retry after a wallpaper is changed! It's soon!");
            return 0;
        }
        CloseHandle(hMutexProcessing);

        if (-1 == updatePictureList(wDirPicture, picList, sizePicList)) {
            HeapFree(GetProcessHeap(), 0, picList);
            return -1;
        }
        if (-1 == processWallpaper(picList, sizePicList, true, emptySpaceColor, upscaleMode, bUseJPEGFormat)) {
            HeapFree(GetProcessHeap(), 0, picList);
            return -1;
        }

        // set wallpaper and exit
        SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, (void*)wWallpaperPath, SPIF_UPDATEINIFILE);
        HeapFree(GetProcessHeap(), 0, picList);
        return 0;
    }

    // wait 1 minute for system boot complete if this instance run at startup
    if (bBootCmd) Sleep(60000);

    // check enough time passed from lastChanged
    {
        FILETIME ft;
        GetSystemTimeAsFileTime(&ft);
        const ULONGLONG sysTime = (((ULONGLONG)ft.dwHighDateTime << 32) + ft.dwLowDateTime) / 10000ULL;
        const long long period = periodInMinute * 60000LL;
        const long long leftTimeToUpdate = period - (long long)(sysTime - getLastChangedReg());

        if (leftTimeToUpdate > 0 && leftTimeToUpdate <= period)
            Sleep((DWORD)leftTimeToUpdate);
    }

    while (true) {
        // if system is busy, wait 1 minute.
        do {
            int busyState = isSystemBusy();
            if (-1 == busyState) {
                HeapFree(GetProcessHeap(), 0, picList);
                CloseHandle(hMutexRunning);
                return -1;
            }

            if (BUSYSTATE_NOT_BUSY == busyState) break;
            
            Sleep(60000UL - BUSYSTATE_CPU_USAGE_CHK_INTVAL * 1000UL);
        } while (true);

        // check FitWallpaper::Processing mutex existance
        // to wait image processing is done in another one-time-instance
        HANDLE hMutexProcessing = NULL;
        while (true) {
            hMutexProcessing = CreateMutexA(NULL, TRUE, "FitWallpaper::Processing");
            if (NULL == hMutexProcessing) {
                DisplayErrorBoxW(L"Failed to create mutex!");
                HeapFree(GetProcessHeap(), 0, picList);
                CloseHandle(hMutexRunning);
                return -1;
            }
            if (GetLastError() == ERROR_ALREADY_EXISTS) {
                CloseHandle(hMutexProcessing);
                Sleep(3000UL);
            }
            else {
                break;
            }
        }

        // update picList and process wallpaper file
        const int bListChanged = updatePictureList(wDirPicture, picList, sizePicList);
        if (-1 == bListChanged) break;

        bool bPicListChanged = false;
        if (PIC_LIST_CHANGED == bListChanged) bPicListChanged = true;

        const int bWallpaperProcessed = processWallpaper(picList, sizePicList, bPicListChanged, emptySpaceColor, upscaleMode, bUseJPEGFormat);
        if (-1 == bWallpaperProcessed) break;

        // close handle which has connection with FitWallpaper::Processing mutex
        CloseHandle(hMutexProcessing);

        if (PROC_WALLPAPER_DONE == bWallpaperProcessed) {
            // set wallpaper
            SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, (void*)wWallpaperPath, SPIF_UPDATEINIFILE);
            // update lastChanged value in registry
            if (-1 == updateLastChangedReg()) break;
        }

        // wait until next update time
        Sleep(periodInMinute * 60000UL - BUSYSTATE_CPU_USAGE_CHK_INTVAL * 1000UL);
    }

    // when program reached here, there were something wrong
    // free heap when exit for safety
    HeapFree(GetProcessHeap(), 0, picList);
    // close handle which has connection with mutex when exit for safety
    CloseHandle(hMutexRunning);
    return -1;
} // wWinMain

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// return BUSYSTATE_NOT_BUSY (not busy) or BUSYSTATE_BUSY (busy)
// return -1 (error)
int isSystemBusy() {
    // on vista or later, check UserNotificationState first.
    if (IsWindowsVistaOrGreater()) {
        QUERY_USER_NOTIFICATION_STATE userNotiState = QUNS_NOT_PRESENT;
        if (S_OK != SHQueryUserNotificationState(&userNotiState)) {
            DisplayErrorBoxW(L"SHQueryUserNotificationState failed!");
            return -1;
        }

        // assume that system is busy when it's in fullscreen or PRESENTATION mode, too.
        if (userNotiState == QUNS_BUSY ||
            userNotiState == QUNS_RUNNING_D3D_FULL_SCREEN ||
            userNotiState == QUNS_PRESENTATION_MODE)
            return BUSYSTATE_BUSY;
    }

    // get cpu usage stat before
    FILETIME ftIdle0, ftKernel0, ftUser0;
    if (0 == GetSystemTimes(&ftIdle0, &ftKernel0, &ftUser0)) {
        DisplayErrorBoxW(L"GetSystemTimes failed!");
        return -1;
    }
    const ULONGLONG idle0 = ((ULONGLONG)ftIdle0.dwHighDateTime << 32) + ftIdle0.dwLowDateTime;
    const ULONGLONG kernel0 = ((ULONGLONG)ftKernel0.dwHighDateTime << 32) + ftKernel0.dwLowDateTime;
    const ULONGLONG user0 = ((ULONGLONG)ftUser0.dwHighDateTime << 32) + ftUser0.dwLowDateTime;
    
    // wait BUSYSTATE_CPU_USAGE_CHK_INTVAL seconds
    Sleep(BUSYSTATE_CPU_USAGE_CHK_INTVAL * 1000);

    // get cpu usage stat after
    FILETIME ftIdle1, ftKernel1, ftUser1;
    if (0 == GetSystemTimes(&ftIdle1, &ftKernel1, &ftUser1)) {
        DisplayErrorBoxW(L"GetSystemTimes failed!");
        return -1;
    }
    const ULONGLONG idle1 = ((ULONGLONG)ftIdle1.dwHighDateTime << 32) + ftIdle1.dwLowDateTime;
    const ULONGLONG kernel1 = ((ULONGLONG)ftKernel1.dwHighDateTime << 32) + ftKernel1.dwLowDateTime;
    const ULONGLONG user1 = ((ULONGLONG)ftUser1.dwHighDateTime << 32) + ftUser1.dwLowDateTime;

    const ULONGLONG diffIdle = idle1 - idle0;
    const ULONGLONG diffKernel = kernel1 - kernel0;
    const ULONGLONG diffUser = user1 - user0;

    // assume that system is busy if cpu usage is over
    // BUSY_CPU_USAGE_THRESHOLD(%) for BUSYSTATE_CPU_USAGE_CHK_INTVAL seconds
    if (BUSYSTATE_CPU_USAGE_THRESHOLD < ((diffKernel + diffUser - diffIdle) * 100ULL / (diffKernel + diffUser))) {
        return BUSYSTATE_BUSY;
    }

    // system is not busy
    return BUSYSTATE_NOT_BUSY;
} // isSystemBusy

// check picture directory is modified
// if so update picList and sizePicList data
// return 0 (ok) or -1 (error)
// return 1 (no need to update)
int updatePictureList(const wchar_t* wDirPicture, wchar_t* picList, int &sizePicList) {
    if (nullptr == wDirPicture || nullptr == picList) {
        DisplayErrorBoxW(L"A wDirPicture or picList is null! (updatePictureList)");
        return -1;
    }
    if (-1 == checkDirectory(wDirPicture, L"picture directory")) return -1;

    const size_t lenWDirPicture = wcsnlen_s(wDirPicture, MAX_PATH);
    if (lenWDirPicture > (MAX_PATH - 3)) {
        DisplayErrorBoxW(L"A picture directory path's length is too long!");
        return -1;
    }

    // check picture directory is modified
    HANDLE hDir = CreateFileW(wDirPicture, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, NULL);
    if (hDir == INVALID_HANDLE_VALUE) {
        DisplayErrorBoxW(L"A picture directory is invalid! (CreateFile)");
        return -1;
    }

    static FILETIME ftWriteBefore = { 0 };
    FILETIME ftWrite = { 0 };

    if (0 == GetFileTime(hDir, NULL, NULL, &ftWrite)) {
        DisplayErrorBoxW(L"Failed to get lastWrite time of a picture directory! (GetFileTime)");
        return -1;
    }

    CloseHandle(hDir);

    // compare modified time before after
    if (ftWriteBefore.dwHighDateTime == ftWrite.dwHighDateTime
        && ftWriteBefore.dwLowDateTime == ftWrite.dwLowDateTime)
        return PIC_LIST_NOT_CHANGED;

    // backup changed modified time
    ftWriteBefore = ftWrite;

    // make find query
    wchar_t wszQuery[MAX_PATH] = L"";
    wcscat_s(wszQuery, wDirPicture);
    wcscat_s(wszQuery, L"\\*");

    WIN32_FIND_DATAW wfd;
    HANDLE hFind = INVALID_HANDLE_VALUE;

    hFind = FindFirstFileW(wszQuery, &wfd);
    if (INVALID_HANDLE_VALUE == hFind) {
        DisplayErrorBoxW(L"Failed to get handle from FindFirstFileW!");
        return -1;
    }

    wchar_t* pPicList = picList;
    sizePicList = 0;

    do {
            // check attributes
        if (0 != (wfd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) ||
            0 != (wfd.dwFileAttributes & FILE_ATTRIBUTE_SYSTEM) ||
            // check file extension
            0 != isSupportedImageFile(wfd.cFileName) ||
            // check filesize
            MAX_PICTURE_FILESIZE < (((ULONGLONG)wfd.nFileSizeHigh << 32) + wfd.nFileSizeLow) ||
            // check limit of pictures
            sizePicList >= MAX_PICTURE_NUMBER)
            continue;

        const size_t lenWPicFound = wcsnlen_s(wfd.cFileName, MAX_PATH);

        if (MAX_PATH > lenWPicFound && (lenWDirPicture + 1 + lenWPicFound) < MAX_PATH) {
            memset(pPicList, 0, MAX_PATH * sizeof(wchar_t));
            wcscat_s(pPicList, MAX_PATH, wDirPicture);
            wcscat_s(pPicList, MAX_PATH, L"\\");
            wcscat_s(pPicList, MAX_PATH, wfd.cFileName);

            pPicList += MAX_PATH;
            sizePicList++;
        }
    } while (FindNextFileW(hFind, &wfd) != 0);

    DWORD dwError = GetLastError();
    FindClose(hFind);

    if (dwError != ERROR_NO_MORE_FILES) {
        DisplayErrorBoxW(L"An error is occured at FindNextFileW!");
        return -1;
    }

    // image file not found
    if (sizePicList == 0) {
        DisplayErrorBoxW(L"No image file found from picture directory!");
        return -1;
    }

    return PIC_LIST_CHANGED;
} // updatePictureList

// return 0 (ok) or -1 (error)
int processWallpaper(const wchar_t* picList, const int sizePicList, const bool bPicListChanged, const int emptySpaceColor, const int upscaleMode, const bool bUseJPEGFormat) {
    if (emptySpaceColor < MIN_EMPTY_SPACE_COLOR || emptySpaceColor > MAX_EMPTY_SPACE_COLOR) {
        DisplayErrorBoxW(L"An emptySpaceColor value is incorrect!");
        return -1;
    }
    if (nullptr == picList || sizePicList <= 0 || sizePicList > MAX_PICTURE_NUMBER) {
        DisplayErrorBoxW(L"A picList is incorrect!");
        return -1;
    }

    static int lastIdxSelected = -1;
    int idxSelected = -1;
    if (bPicListChanged) lastIdxSelected = -1;
    // if picList data is not changed and sizePicList is only 1
    // there are no need to update wallpaper
    // because wallpaper which is only one is applied already
    else if (1 == sizePicList) {
        return PROC_WALLPAPER_DUP_PICTURE;
    }

    Mat input = Mat();
    {
        FILE* fp = nullptr;
        long long sizePicture = -1;

        while (true) {
            unsigned int rndNumber = -1;
            errno_t err = rand_s(&rndNumber);
            if (err != 0) {
                DisplayErrorBoxW(L"Failed to call rand_s function!");
                return -1;
            }

            idxSelected = int((double)rndNumber / ((double)UINT_MAX + 1) * sizePicList);

            if (idxSelected != lastIdxSelected) break;
        }

        const wchar_t* pPath = picList;
        pPath += idxSelected * MAX_PATH;

        _wfopen_s(&fp, pPath, L"rb");
        if (nullptr == fp) {
            DisplayErrorBoxW(L"Failed to open picture file!");
            return -1;
        }
        const int fileno = _fileno(fp);
        if (-1 == fileno) {
            fclose(fp);
            DisplayErrorBoxW(L"Failed to get picture file's descriptor!");
            return -1;
        }
        sizePicture = _filelengthi64(fileno);
        if (-1LL == sizePicture) {
            fclose(fp);
            DisplayErrorBoxW(L"Failed to get picture file's size!");
            return -1;
        }

        vector<BYTE> picData(sizePicture);
        const size_t rdBytes = fread(&picData[0], sizeof(BYTE), sizePicture, fp);
        if (rdBytes != sizePicture) {
            fclose(fp);
            DisplayErrorBoxW(L"Failed to read picture file!");
            return -1;
        }

        input = imdecode(move(picData), IMREAD_UNCHANGED);
        fclose(fp);
    }

    if (input.empty()) {
        DisplayErrorBoxW(L"Failed to read image file!");
        return -1;
    }

    // convert to 8 bits color channel
    normalize(input, input, 0, UCHAR_MAX, NORM_MINMAX, CV_8U);

    // save channel count of input
    const int inputChannels = input.channels();

    // shrink if need
    const int desktop_x = GetSystemMetrics(SM_CXSCREEN);
    const int desktop_y = GetSystemMetrics(SM_CYSCREEN);
    const double sx = (double)desktop_x / input.cols;
    const double sy = (double)desktop_y / input.rows;

    // shrink only if scale value < 1.0
    if (sx > sy && sy < 1.0)
        resize(input, input, Size(), sy, sy, INTER_CUBIC);
    else if (sx <= sy && sx < 1.0)
        resize(input, input, Size(), sx, sx, INTER_CUBIC);

    // fallback empty space color is white
    int esColorB = 255, esColorG = 255, esColorR = 255;

    // if empty space exist or image has alpha (transparancy) data (BGRA)
    // update empty space color
    if (input.cols < desktop_x || input.rows < desktop_y || inputChannels == 4) {
        if (inputChannels == 4) {
            for (int i = 0; i < input.rows; i++) {
                for (int j = 0; j < input.cols; j++)
                {
                    Vec4b& v = input.at<Vec4b>(i, j);
                    if (v[3] == 0) v = Scalar(0, 0, 0, 0);
                }
            }
        }

        if (emptySpaceColor == EMPTY_SPACE_COLOR_D) {
            Mat m = input.reshape(1, input.rows * input.cols);
            m.convertTo(m, CV_32F);
            Mat labels, centers;

            kmeans(m, 6, labels, TermCriteria(TermCriteria::COUNT | TermCriteria::EPS, 10, 1.0),
                1, KMEANS_RANDOM_CENTERS, centers);

            int hist[6] = {};
            for (int i = 0; i < labels.rows; i++)
                hist[labels.at<int>(i, 0)]++;

            int maxIdx = -1;

            if (inputChannels == 4) {
                int max = 0;
                for (int i = 0; i < 6; i++) {
                    if (int(centers.at<float>(i, 3)) >= 192 && hist[i] > max) {
                        max = hist[i];
                        maxIdx = i;
                    }
                }

                if (maxIdx != -1) {
                    esColorB = int(centers.at<float>(maxIdx, 0));
                    esColorG = int(centers.at<float>(maxIdx, 1));
                    esColorR = int(centers.at<float>(maxIdx, 2));
                }
            }
            else {
                maxIdx = int(distance(hist, max_element(hist, hist + 6)));
                esColorB = int(centers.at<float>(maxIdx, 0));
                if (inputChannels == 3) {
                    esColorG = int(centers.at<float>(maxIdx, 1));
                    esColorR = int(centers.at<float>(maxIdx, 2));
                }
                else if (inputChannels == 1) {
                    esColorG = esColorB;
                    esColorR = esColorB;
                }
            }
        }
        else if (emptySpaceColor == EMPTY_SPACE_COLOR_B) {
            esColorB = 0;
            esColorG = 0;
            esColorR = 0;
        }

        if (inputChannels == 4) {
            for (int i = 0; i < input.rows; i++) {
                for (int j = 0; j < input.cols; j++)
                {
                    Vec4b& v = input.at<Vec4b>(i, j);
                    if (v[3] != 255) {
                        v[0] = esColorB + (v[0] - esColorB) * v[3] / 255;
                        v[1] = esColorG + (v[1] - esColorG) * v[3] / 255;
                        v[2] = esColorR + (v[2] - esColorR) * v[3] / 255;
                    }
                }
            }
            cvtColor(input, input, COLOR_BGRA2BGR);
        }
    }

    // upscaleMode check
    if (UPSCALE_MODE_DONT_UPSCALE < upscaleMode) {
        // upscale if need
        const double ux = (double)desktop_x / input.cols;
        const double uy = (double)desktop_y / input.rows;
        double upscale = 1.0;

        // upscale only if scale value > 1.0
        if (ux > uy && uy > 1.0) upscale = uy;
        else if (ux <= uy && ux > 1.0) upscale = ux;

        // To do upscale, scale factor must be bigger than 1.0
        if (upscale > 1.0) {
            // limit maximum scale by 2 if 2X option used
            if (upscale > 2.0 && UPSCALE_MODE_UPSCALE_UP_TO_2X == upscaleMode)
                upscale = 2.0;
            // limit maximum scale by 4 if 4X option used
            if (upscale > 4.0 && UPSCALE_MODE_UPSCALE_UP_TO_4X == upscaleMode)
                upscale = 4.0;

            resize(input, input, Size(), upscale, upscale, INTER_CUBIC);
        }
    }

    // if empty space exist
    if (input.cols < desktop_x || input.rows < desktop_y) {
        int bTop = 0, bLeft = 0, bBottom = 0, bRight = 0;

        if (input.cols < desktop_x) {
            int diff = desktop_x - input.cols;
            bLeft = bRight = diff / 2;
            if (diff % 2 == 1) bLeft++;
        }
        if (input.rows < desktop_y) {
            int diff = desktop_y - input.rows;
            bTop = bBottom = diff / 2;
            if (diff % 2 == 1) bBottom++;
        }

        copyMakeBorder(input, input,
            bTop, bBottom, bLeft, bRight,
            BORDER_CONSTANT, Scalar(esColorB, esColorG, esColorR));
    }

    int ret = -1;
    if (bUseJPEGFormat) {
        vector<int> params;
        params.push_back(IMWRITE_JPEG_QUALITY);
        params.push_back(100);

        ret = imwrite(STR_WALLPAPER_FNAME_JPEG, input, params);
    }
    else
        ret = imwrite(STR_WALLPAPER_FNAME_PNG, input);

    if (ret) {
        // save last index of picture used when success to write output
        lastIdxSelected = idxSelected;
        return PROC_WALLPAPER_DONE;
    }
    else {
        DisplayErrorBoxW(L"Failed to write wallpaper file!");
        return -1;
    }
} // processWallpaper

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

void DisplayInfoBoxA(LPCSTR szInfoMsg) {
    if (nullptr == szInfoMsg)
        MessageBoxA(NULL, "A szInfoMsg is null! (DisplayInfoBoxA)", "FitWallpaper: Info", MB_ICONINFORMATION | MB_OK);
    else
        MessageBoxA(NULL, szInfoMsg, "FitWallpaper: Info", MB_ICONINFORMATION | MB_OK);
} // DisplayInfoBoxA

void DisplayErrorBoxW(LPCWSTR wszErrorMsg) {
    if (nullptr == wszErrorMsg) {
        MessageBoxW(NULL, L"A wszErrorMsg is null! (DisplayErrorBoxW)", L"FitWallpaper: Error", MB_ICONERROR | MB_OK);
        return;
    }

    LPWSTR pMsgBuf = nullptr;
    LPWSTR pDisplayBuf;
    DWORD dw = GetLastError();
    wchar_t wszErrCode[16] = L"";
    
    if (0 != _ultow_s(dw, wszErrCode, 10)) {
        MessageBoxW(NULL, wszErrorMsg, L"FitWallpaper: Error", MB_ICONERROR | MB_OK);
        return;
    }

    if (!FormatMessageW(
        FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
        NULL, dw, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
        (LPWSTR)&pMsgBuf, 0, NULL))
    {
        MessageBoxW(NULL, wszErrorMsg, L"FitWallpaper: Error", MB_ICONERROR | MB_OK);
        return;
    }

    HANDLE hDefaultHeap = GetProcessHeap();
    if (nullptr == hDefaultHeap) {
        MessageBoxW(NULL, wszErrorMsg, L"FitWallpaper: Error", MB_ICONERROR | MB_OK);
        return;
    }

    const size_t wLength = wcslen(pMsgBuf) + wcslen(wszErrorMsg) + 16;
    pDisplayBuf = (LPWSTR)HeapAlloc(hDefaultHeap, HEAP_ZERO_MEMORY, wLength * sizeof(WCHAR));
    if (nullptr == pDisplayBuf) {
        LocalFree(pMsgBuf);
        MessageBoxW(NULL, wszErrorMsg, L"FitWallpaper: Error", MB_ICONERROR | MB_OK);
        return;
    }

    wcscat_s(pDisplayBuf, wLength, wszErrorMsg);
    wcscat_s(pDisplayBuf, wLength, L"\n");
    wcscat_s(pDisplayBuf, wLength, wszErrCode);
    wcscat_s(pDisplayBuf, wLength, L": ");
    wcscat_s(pDisplayBuf, wLength, pMsgBuf);

    MessageBoxW(NULL, pDisplayBuf, L"FitWallpaper: Error", MB_ICONERROR | MB_OK);

    LocalFree(pMsgBuf);
    HeapFree(hDefaultHeap, 0, pDisplayBuf);
} // DisplayErrorBoxW

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// return 0 (ok) or -1 (error)
int registerRunAtStartupReg(const wchar_t* wszDirWorking) {
    if (nullptr == wszDirWorking) {
        DisplayErrorBoxW(L"A wszDirWorking is null! (registerRunAtStartupReg)");
        return -1;
    }
    if (-1 == checkDirectory(wszDirWorking, L"current directory")) return -1;

    HKEY hKey;
    const wchar_t* wszValueName = L"FitWallpaper";

    LSTATUS lResult = RegOpenKeyExW(HKEY_CURRENT_USER,
        L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run",
        0, KEY_SET_VALUE, &hKey);

    if (ERROR_SUCCESS == lResult) {
        const size_t wLen = wcsnlen_s(wszDirWorking, MAX_PATH);
        if (wLen >= (MAX_PATH - wcslen(WSTR_PROGRAM_NAME) - 1 - wcslen(WSTR_CMDLINE_BOOT) - 1 - 1)) {
            DisplayErrorBoxW(L"A wszDirWorking is too long! (registerRunAtStartupReg)");
            return -1;
        }

        wchar_t wszValueData[MAX_PATH] = L"";
        wcscat_s(wszValueData, wszDirWorking);
        wcscat_s(wszValueData, L"\\");
        wcscat_s(wszValueData, WSTR_PROGRAM_NAME);
        wcscat_s(wszValueData, L" ");
        wcscat_s(wszValueData, WSTR_CMDLINE_BOOT);

        lResult = RegSetValueExW(hKey,
            wszValueName,
            0,
            REG_SZ,
            (BYTE*)wszValueData,
            (DWORD)(sizeof(wchar_t) * (wcslen(wszValueData) + 1ULL)));

        RegCloseKey(hKey);

        if (ERROR_SUCCESS != lResult) {
            DisplayErrorBoxW(L"Failed to write FitWallpaper value into registry! (registerRunAtStartupReg)");
            return -1;
        }
    }
    else {
        DisplayErrorBoxW(L"Failed to retrive Run key from registry! (registerRunAtStartupReg)");
        return -1;
    }

    return 0;
} // registerRunAtStartupReg

// return 0 (ok) or -1 (error)
int unregisterRunAtStartupReg() {
    HKEY hKey;
    const wchar_t* wszValueName = L"FitWallpaper";

    LSTATUS lResult = RegOpenKeyExW(HKEY_CURRENT_USER,
        L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run",
        0, KEY_SET_VALUE, &hKey);

    if (ERROR_SUCCESS == lResult) {
        lResult = RegDeleteValueW(hKey, wszValueName);

        RegCloseKey(hKey);

        if (ERROR_SUCCESS != lResult && ERROR_FILE_NOT_FOUND != lResult) {
            DisplayErrorBoxW(L"Failed to delete FitWallpaper value from registry! (unregisterRunAtStartupReg)");
            return -1;
        }
    }
    else {
        DisplayErrorBoxW(L"Failed to retrive Run key from registry! (unregisterRunAtStartupReg)");
        return -1;
    }

    return 0;
} // unregisterRunAtStartupReg

// return 0 (ok) or -1 (error)
int updateJPEGImportQualityReg() {
    HKEY hKey;
    DWORD regType = REG_DWORD;
    DWORD cbSize = sizeof(DWORD);
    DWORD value = 0;

    LSTATUS lResult = RegOpenKeyExW(HKEY_CURRENT_USER,
        L"Control Panel\\Desktop",
        0, KEY_QUERY_VALUE | KEY_SET_VALUE, &hKey);

    if (ERROR_SUCCESS == lResult) {
        lResult = RegQueryValueExW(hKey,
            L"JPEGImportQuality",
            NULL,
            &regType,
            (BYTE*)&value,
            &cbSize);

        if (ERROR_SUCCESS != lResult && ERROR_FILE_NOT_FOUND != lResult) {
            DisplayErrorBoxW(L"Failed to query JPEGImportQuality value from registry! (updateJPEGImportQualityReg)");
            return -1;
        }

        if (value < 100 || ERROR_FILE_NOT_FOUND == lResult) {
            value = 100;
            lResult = RegSetValueExW(hKey,
                L"JPEGImportQuality",
                0,
                REG_DWORD,
                (BYTE*)&value,
                (DWORD)sizeof(DWORD));

            RegCloseKey(hKey);

            if (ERROR_SUCCESS != lResult) {
                DisplayErrorBoxW(L"Failed to write JPEGImportQuality value into registry! (updateJPEGImportQualityReg)");
                return -1;
            }
            else if (!IsWindows8OrGreater()) {
                DisplayInfoBoxA("Please, restart your computer for take effect of JPEGImportQuality registry");
            }
        }
    }
    else {
        DisplayErrorBoxW(L"Failed to retrive Desktop key from registry! (updateJPEGImportQualityReg)");
        return -1;
    }

    return 0;
} // updateJPEGImportQualityReg


// return last changed time in ms from registry
// if lastChanged value not found or error, return 0ULL.
ULONGLONG getLastChangedReg() {
    HKEY hKey;
    const wchar_t* wszValueName = L"lastChanged";
    ULONGLONG lastChanged = 0ULL;
    DWORD regType = REG_BINARY;
    DWORD cbSize = sizeof(ULONGLONG);

    LSTATUS lResult = RegOpenKeyExW(HKEY_CURRENT_USER,
        L"SOFTWARE\\FitWallpaper", 0, KEY_QUERY_VALUE, &hKey);

    if (ERROR_SUCCESS == lResult) {
        lResult = RegQueryValueExW(hKey,
            wszValueName,
            NULL,
            &regType,
            (BYTE*)&lastChanged,
            &cbSize);

        RegCloseKey(hKey);

        if (ERROR_SUCCESS != lResult && ERROR_FILE_NOT_FOUND != lResult) {
            DisplayErrorBoxW(L"Failed to query lastChanged value from registry! (getLastChangedReg)");
        }
    }
    else if (ERROR_FILE_NOT_FOUND != lResult) {
        DisplayErrorBoxW(L"Failed to retrive key from registry! (getLastChangedReg)");
    }

    return lastChanged;
} // getLastChangedReg

// return 0 (ok) or -1 (error)
// write a current time in ms as lastChanged value into registry
int updateLastChangedReg() {
    FILETIME ft;
    GetSystemTimeAsFileTime(&ft);

    ULONGLONG sysTime = (((ULONGLONG)ft.dwHighDateTime << 32) + ft.dwLowDateTime) / 10000ULL;

    HKEY hKey;
    const wchar_t* wszValueName = L"lastChanged";

    LSTATUS lResult = RegCreateKeyExW(HKEY_CURRENT_USER,
        L"SOFTWARE\\FitWallpaper",
        NULL, NULL,
        REG_OPTION_NON_VOLATILE,
        KEY_WRITE,
        NULL,
        &hKey,
        NULL);

    if (ERROR_SUCCESS == lResult) {
        lResult = RegSetValueExW(hKey,
            wszValueName,
            0,
            REG_BINARY,
            (BYTE*)&sysTime,
            (DWORD)(sizeof(ULONGLONG)));

        RegCloseKey(hKey);

        if (ERROR_SUCCESS != lResult) {
            DisplayErrorBoxW(L"Failed to write lastChanged value into registry! (updateLastChangedReg)");
            return -1;
        }
    }
    else {
        DisplayErrorBoxW(L"Failed to create or retrive key from registry! (updateLastChangedReg)");
        return -1;
    }

    return 0;
} // updateLastChangedReg

// return 0 (ok) or -1 (error)
int unregisterKeyFromReg() {
    LSTATUS lResult = RegDeleteKeyW(HKEY_CURRENT_USER, L"SOFTWARE\\FitWallpaper");

    if (ERROR_SUCCESS != lResult && ERROR_FILE_NOT_FOUND != lResult) {
        DisplayErrorBoxW(L"Failed to delete entire key from registry! (unregisterKeyFromReg)");
        return -1;
    }

    return 0;
} // unregisterKeyFromReg

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// return 0 (ok) or -1 (error)
int checkDirectory(const wchar_t* dir, const wchar_t* dirName) {
    if (nullptr == dir || nullptr == dirName) {
        DisplayErrorBoxW(L"A dir or dirName is null! (checkDirectory)");
        return -1;
    }
    if (-1 == isIllegalCharacterExist(dir, MAX_PATH)) {
        DisplayErrorBoxW(L"A dir has illegal character(s)! (checkDirectory)");
        return -1;
    }
    if (wcsnlen_s(dirName, 74) == 74) {
        DisplayErrorBoxW(L"A dirName's length is too long! (checkDirectory)");
        return -1;
    }

    wchar_t msg[100] = L"";
    wcscat_s(msg, L"A ");
    wcscat_s(msg, dirName);

    DWORD dwAttrib = GetFileAttributesW(dir);
    if (dwAttrib == INVALID_FILE_ATTRIBUTES) {
        wcscat_s(msg, L" is invalid!");
        DisplayErrorBoxW(msg);
        return -1;
    }
    else if ((dwAttrib & FILE_ATTRIBUTE_SYSTEM) != 0) {
        wcscat_s(msg, L" is a system directory!");
        DisplayErrorBoxW(msg);
        return -1;
    }
    else if ((dwAttrib & FILE_ATTRIBUTE_DIRECTORY) == 0) {
        wcscat_s(msg, L" is a file!");
        DisplayErrorBoxW(msg);
        return -1;
    }
    return 0;
} // checkDirectory

// return 0 (ok) or -1 (error)
int checkFile(const wchar_t* file, const wchar_t* fileName) {
    if (nullptr == file || nullptr == fileName) {
        DisplayErrorBoxW(L"A file or fileName is null! (checkFile)");
        return -1;
    }
    if (-1 == isIllegalCharacterExist(file, MAX_PATH)) {
        DisplayErrorBoxW(L"A file has illegal character(s)! (checkFile)");
        return -1;
    }
    if (wcsnlen_s(fileName, 79) == 79) {
        DisplayErrorBoxW(L"A fileName's length is too long! (checkFile)");
        return -1;
    }

    wchar_t msg[100] = L"";
    wcscat_s(msg, L"A ");
    wcscat_s(msg, fileName);

    DWORD dwAttrib = GetFileAttributesW(file);
    if (dwAttrib == INVALID_FILE_ATTRIBUTES) {
        wcscat_s(msg, L" is invalid!");
        DisplayErrorBoxW(msg);
        return -1;
    }
    else if ((dwAttrib & FILE_ATTRIBUTE_SYSTEM) != 0) {
        wcscat_s(msg, L" is a system file!");
        DisplayErrorBoxW(msg);
        return -1;
    }
    else if ((dwAttrib & FILE_ATTRIBUTE_DIRECTORY) != 0) {
        wcscat_s(msg, L" is a dirctory!");
        DisplayErrorBoxW(msg);
        return -1;
    }
    return 0;
} // checkFile

// return 0 (ok and supported image) or -1 (error or not supported image or file)
int isSupportedImageFile(const wchar_t* imageFileName) {
    if (nullptr == imageFileName) {
        DisplayErrorBoxW(L"An imageFileName is null! (isSupportedImageFile)");
        return -1;
    }
    if (-1 == isIllegalCharacterExist(imageFileName, MAX_PATH)) {
        DisplayErrorBoxW(L"An imageFileName has illegal character(s)! (isSupportedImageFile)");
        return -1;
    }

    const wchar_t* pExt = imageFileName;
    const size_t wstrLen = wcsnlen_s(imageFileName, MAX_PATH);

    if (MAX_PATH > wstrLen && 3 < wstrLen) {
        int rdCount = 0;
        pExt += wstrLen; pExt--;
        while (pExt >= imageFileName && *pExt != '.') {
            pExt--; rdCount++;
            if (rdCount == 4) break;
        }

        if (pExt >= imageFileName && *pExt == '.') {
            pExt++;

            for (const wchar_t* szImageExt : WSTR_IMAGE_EXT_LIST) {
                if (0 == _wcsnicmp(pExt, szImageExt, wcsnlen_s(szImageExt, MAX_IMAGE_EXT_LENGTH + 1)))
                    return 0;
            }
        }
    }

    DisplayErrorBoxW(L"An imageFileName is not supported image file! (isSupportedImageFile)");
    return -1;
} // isSupportedImageFile

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// return 0 (ok) or -1 (error)
int cvt_utf8_to_wchar(const char* src, const int srcSize, wchar_t* dest, const int destSize) {
    if (src != nullptr && srcSize > 0 && dest != nullptr && destSize > 0) {
        if (260 == strnlen_s(src, srcSize)) {
            DisplayErrorBoxW(L"A src's length is too long! (cvt_utf8_to_wchar)");
            return -1;
        }

        const int bufSizeRequired = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, src, -1, NULL, 0);
        if (bufSizeRequired <= 0) {
            DisplayErrorBoxW(L"A src is invalid or it may contain illegal utf8 character(s)! (cvt_utf8_to_wchar)");
            return -1;
        }
        else if (bufSizeRequired > destSize) {
            DisplayErrorBoxW(L"A src's length is too long to convert it to wchar! (cvt_utf8_to_wchar)");
            return -1;
        }
        else {
            const int sizeConverted = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, src, -1, dest, destSize);
            if (sizeConverted <= 0) {
                DisplayErrorBoxW(L"Failed to convert src to wchar! (cvt_utf8_to_wchar)");
                return -1;
            }
        }
    }
    else {
        DisplayErrorBoxW(L"A src or dest or both is invalid! (cvt_utf8_to_wchar)");
        return -1;
    }

    return 0;
} // cvt_utf8_to_wchar

// return 0 (ok) or -1 (error)
int cvt_wchar_to_utf8(const wchar_t* src, const int srcSize, char* dest, const int destSize) {
    if (src != nullptr && srcSize > 0 && dest != nullptr && destSize > 0) {
        if (260 == wcsnlen_s(src, srcSize)) {
            DisplayErrorBoxW(L"A src's length is too long! (cvt_wchar_to_utf8)");
            return -1;
        }

        // WC_ERR_INVALID_CHARS is only allowed in Windows Vista or later
        // in Windows XP, we have to check if U+FFFD(utf8: EF BF BD { -17, -65, -67, 0 }) value exist in dest
        DWORD dwFlag = 0;
        if (IsWindowsVistaOrGreater()) {
            dwFlag = WC_ERR_INVALID_CHARS;
        }

        const int bufSizeRequired = WideCharToMultiByte(CP_UTF8, dwFlag, src, -1, NULL, 0, NULL, NULL);
        if (bufSizeRequired <= 0) {
            DisplayErrorBoxW(L"A src is invalid or it may contain illegal wchar character(s)! (cvt_wchar_to_utf8)");
            return -1;
        }
        else if (bufSizeRequired > destSize) {
            DisplayErrorBoxW(L"A src's length is too long to convert it to utf8! (cvt_wchar_to_utf8)");
            return -1;
        }
        else {
            const int sizeConverted = WideCharToMultiByte(CP_UTF8, dwFlag, src, -1, dest, destSize, NULL, NULL);
            if (sizeConverted <= 0) {
                DisplayErrorBoxW(L"Failed to convert src to utf8! (cvt_wchar_to_utf8)");
                return -1;
            }
            // (Windows XP) check if U + FFFD(utf8: EF BF BD{ -17, -65, -67, 0 }) value exist in dest
            else if (0 == dwFlag) {
                const char illegalChar[4] = { -17, -65, -67, 0 };
                const char* pIllegal = strstr(dest, illegalChar);
                if (nullptr != pIllegal) {
                    DisplayErrorBoxW(L"A src contain invalid wchar character(s)! (cvt_wchar_to_utf8)");
                    return -1;
                }
            }
        }
    }
    else {
        DisplayErrorBoxW(L"A src or dest or both is invalid! (cvt_wchar_to_utf8)");
        return -1;
    }

    return 0;
} // cvt_wchar_to_utf8

// return 0 (ok) or -1 (error)
int isIllegalCharacterExist(const wchar_t* wszDir, const int size) {
    if (nullptr == wszDir || size <= 0) return -1;

    const wchar_t* pCur = wszDir;
    const wchar_t* pMax = pCur + min(wcsnlen_s(wszDir, size), size);
    
    // illegal character
    // /*?"<>| and tab character(\t)
    while (pCur < pMax && *pCur != L'\0') {
        if (L'/'  == *pCur || L'*'  == *pCur || L'?' == *pCur ||
            L'\"' == *pCur || L'<'  == *pCur || L'>' == *pCur ||
            L'|'  == *pCur || L'\t' == *pCur)
            return -1;
        pCur++;
    }

    return 0;
} // isIllegalCharacterExist