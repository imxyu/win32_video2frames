/*
    Video Frame Extractor Tool (Batch Processing Version - Final Fix)
    Fix: Added NOMINMAX, std::min/max, and cstdint header.
*/

// 1. 强制窗口子系统
#pragma comment(linker, "/SUBSYSTEM:WINDOWS")

// 2. 避免 min/max 宏冲突 (必须在 windows.h 之前)
#define NOMINMAX

// 3. 禁用安全警告
#define _CRT_SECURE_NO_WARNINGS

#include <windows.h>
#include <commctrl.h>
#include <shlwapi.h>
#include <shlobj.h>
#include <gdiplus.h>
#include <mfapi.h>
#include <mfidl.h>
#include <mfreadwrite.h>
#include <propvarutil.h>
#include <string>
#include <vector>
#include <atomic>
#include <thread>
#include <cmath>
#include <algorithm> // 用于 std::min, std::max
#include <cstdint>   // 用于 int8_t 等类型
#include <cstdio>    // 用于 swprintf

// 链接库
#pragma comment(lib, "gdiplus.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "mf.lib")
#pragma comment(lib, "mfplat.lib")
#pragma comment(lib, "mfreadwrite.lib")
#pragma comment(lib, "mfuuid.lib")
#pragma comment(lib, "propsys.lib")

using namespace Gdiplus;
using namespace std;

// 安全释放宏
template <class T> void SafeRelease(T** ppT) {
    if (*ppT) { (*ppT)->Release(); *ppT = NULL; }
}

// 控件ID
#define IDC_EDT_PATH    1001
#define IDC_EDT_OUT     1002
#define IDC_BTN_BROWSE  1003
#define IDC_CHK_AUTO    1004
#define IDC_EDT_INT     1005
#define IDC_EDT_X1      1006
#define IDC_EDT_Y1      1007
#define IDC_EDT_X2      1008
#define IDC_EDT_Y2      1009
#define IDC_BTN_START   1010
#define IDC_PREVIEW     1011
#define IDC_PROGRESS    1012
#define IDC_LBL_INFO    1013
#define IDC_LBL_BATCH   1014
#define IDC_LBL_ROI     1015

// 全局状态
HINSTANCE hInst;
HWND hMainWnd;
ULONG_PTR gdiplusToken;

// 批量处理相关变量
vector<wstring> g_batchFiles;
bool g_isConsistent = true;
UINT32 g_batchWidth = 0;
UINT32 g_batchHeight = 0;

// 预览相关
UINT64 g_durationHns = 0;
double g_durationSec = 0.0;
double g_fps = 0.0;
Bitmap* g_pPreviewBitmap = nullptr;
bool g_isExtracting = false;
std::atomic<bool> g_stopRequested(false);

// ==========================================
// 前置声明
// ==========================================
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK PreviewProc(HWND, UINT, WPARAM, LPARAM, UINT_PTR, DWORD_PTR);
void ProcessDrop(const wstring& path);
void ProcessMultipleFiles(const vector<wstring>& files);
void StartExtractionThread();
int GetEncoderClsid(const WCHAR* format, CLSID* pClsid);
bool BrowseFolder(HWND hWnd, wstring& outPath);
void SetIntToEdit(int id, int val);
int GetIntFromEdit(int id);
void UpdateROISizeLabel();
bool IsVideoFile(const wstring& path);
// ==========================================

// ==========================================
// Media Foundation 视频读取器
// ==========================================
class VideoReaderMF {
public:
    IMFSourceReader* m_pReader;

    VideoReaderMF() : m_pReader(NULL) {}
    ~VideoReaderMF() { Close(); }

    void Close() { SafeRelease(&m_pReader); }

    HRESULT Open(const wstring& filepath) {
        Close();
        IMFAttributes* pAttributes = NULL;
        HRESULT hr = MFCreateAttributes(&pAttributes, 1);
        if (SUCCEEDED(hr)) {
            hr = pAttributes->SetUINT32(MF_SOURCE_READER_ENABLE_VIDEO_PROCESSING, TRUE);
        }
        hr = MFCreateSourceReaderFromURL(filepath.c_str(), pAttributes, &m_pReader);
        SafeRelease(&pAttributes);
        if (FAILED(hr)) return hr;

        hr = m_pReader->SetStreamSelection(MF_SOURCE_READER_ALL_STREAMS, FALSE);
        hr = m_pReader->SetStreamSelection(MF_SOURCE_READER_FIRST_VIDEO_STREAM, TRUE);
        if (FAILED(hr)) return hr;

        IMFMediaType* pType = NULL;
        hr = MFCreateMediaType(&pType);
        if (SUCCEEDED(hr)) {
            pType->SetGUID(MF_MT_MAJOR_TYPE, MFMediaType_Video);
            pType->SetGUID(MF_MT_SUBTYPE, MFVideoFormat_RGB32);
            hr = m_pReader->SetCurrentMediaType(MF_SOURCE_READER_FIRST_VIDEO_STREAM, NULL, pType);
            SafeRelease(&pType);
        }
        return hr;
    }

    HRESULT GetVideoInfo(UINT32& w, UINT32& h, UINT64& duration, double& fps) {
        if (!m_pReader) return E_FAIL;
        
        // 初始化输出参数
        w = 0;
        h = 0;
        duration = 0;
        fps = 0.0;
        
        IMFMediaType* pType = NULL;
        HRESULT hr = m_pReader->GetCurrentMediaType(MF_SOURCE_READER_FIRST_VIDEO_STREAM, &pType);
        if (FAILED(hr)) return hr;
        MFGetAttributeSize(pType, MF_MT_FRAME_SIZE, &w, &h);
        UINT32 num = 0, den = 1;
        if (SUCCEEDED(MFGetAttributeRatio(pType, MF_MT_FRAME_RATE, &num, &den)) && den != 0) {
            fps = (double)num / (double)den;
        }
        SafeRelease(&pType);
        PROPVARIANT var;
        PropVariantInit(&var);
        if (SUCCEEDED(m_pReader->GetPresentationAttribute(MF_SOURCE_READER_MEDIASOURCE, MF_PD_DURATION, &var))) {
            if (var.vt == VT_UI8) duration = var.uhVal.QuadPart;
            PropVariantClear(&var);
        }
        return S_OK;
    }

    HRESULT Seek(double seconds) {
        if (!m_pReader) return E_FAIL;
        PROPVARIANT var;
        PropVariantInit(&var);
        var.vt = VT_I8;
        var.hVal.QuadPart = (LONGLONG)(seconds * 10000000.0);
        HRESULT hr = m_pReader->SetCurrentPosition(GUID_NULL, var);
        PropVariantClear(&var);
        // 注意：不要调用 Flush()，这会导致解码状态异常
        return hr;
    }

    Bitmap* ReadFrame(UINT32 width, UINT32 height) {
        if (!m_pReader) return nullptr;
        IMFSample* pSample = NULL;
        DWORD flags = 0;
        HRESULT hr = m_pReader->ReadSample(MF_SOURCE_READER_FIRST_VIDEO_STREAM, 0, NULL, &flags, NULL, &pSample);
        if (FAILED(hr) || pSample == NULL) return nullptr;

        Bitmap* safeBmp = nullptr;
        IMFMediaBuffer* pBuffer = NULL;
        if (SUCCEEDED(pSample->ConvertToContiguousBuffer(&pBuffer))) {
            BYTE* pSrcData = NULL;
            DWORD srcLen = 0;
            if (SUCCEEDED(pBuffer->Lock(&pSrcData, NULL, &srcLen))) {
                safeBmp = new Bitmap(width, height, PixelFormat32bppRGB);
                if (safeBmp->GetLastStatus() == Ok) {
                    BitmapData bmpData;
                    Rect rect(0, 0, width, height);
                    if (safeBmp->LockBits(&rect, ImageLockModeWrite, PixelFormat32bppRGB, &bmpData) == Ok) {
                        BYTE* pDstRow = (BYTE*)bmpData.Scan0;
                        BYTE* pSrcRow = pSrcData;
                        int bytesPerRow = width * 4;
                        for (UINT32 y = 0; y < height; y++) {
                            if ((pSrcRow - pSrcData) + bytesPerRow > (int)srcLen) break;
                            memcpy(pDstRow, pSrcRow, bytesPerRow);
                            pDstRow += bmpData.Stride;
                            pSrcRow += bytesPerRow;
                        }
                        safeBmp->UnlockBits(&bmpData);
                    }
                    else { delete safeBmp; safeBmp = nullptr; }
                }
                else { delete safeBmp; safeBmp = nullptr; }
                pBuffer->Unlock();
            }
            SafeRelease(&pBuffer);
        }
        SafeRelease(&pSample);
        return safeBmp;
    }

    // 顺序读取下一帧，返回帧的时间戳（单位：100纳秒）
    Bitmap* ReadNextFrame(UINT32 width, UINT32 height, LONGLONG* outTimestamp) {
        if (!m_pReader) return nullptr;
        
        IMFSample* pSample = NULL;
        DWORD flags = 0;
        LONGLONG timestamp = 0;
        
        HRESULT hr = m_pReader->ReadSample(
            MF_SOURCE_READER_FIRST_VIDEO_STREAM, 
            0, 
            NULL, 
            &flags, 
            &timestamp, 
            &pSample
        );
        
        if (FAILED(hr) || pSample == NULL || (flags & MF_SOURCE_READERF_ENDOFSTREAM)) {
            if (pSample) SafeRelease(&pSample);
            return nullptr;
        }
        
        if (outTimestamp) *outTimestamp = timestamp;
        
        Bitmap* safeBmp = nullptr;
        IMFMediaBuffer* pBuffer = NULL;
        if (SUCCEEDED(pSample->ConvertToContiguousBuffer(&pBuffer))) {
            BYTE* pSrcData = NULL;
            DWORD srcLen = 0;
            if (SUCCEEDED(pBuffer->Lock(&pSrcData, NULL, &srcLen))) {
                safeBmp = new Bitmap(width, height, PixelFormat32bppRGB);
                if (safeBmp->GetLastStatus() == Ok) {
                    BitmapData bmpData;
                    Rect rect(0, 0, width, height);
                    if (safeBmp->LockBits(&rect, ImageLockModeWrite, PixelFormat32bppRGB, &bmpData) == Ok) {
                        BYTE* pDstRow = (BYTE*)bmpData.Scan0;
                        BYTE* pSrcRow = pSrcData;
                        int bytesPerRow = width * 4;
                        for (UINT32 y = 0; y < height; y++) {
                            if ((pSrcRow - pSrcData) + bytesPerRow > (int)srcLen) break;
                            memcpy(pDstRow, pSrcRow, bytesPerRow);
                            pDstRow += bmpData.Stride;
                            pSrcRow += bytesPerRow;
                        }
                        safeBmp->UnlockBits(&bmpData);
                    }
                    else { delete safeBmp; safeBmp = nullptr; }
                }
                else { delete safeBmp; safeBmp = nullptr; }
                pBuffer->Unlock();
            }
            SafeRelease(&pBuffer);
        }
        SafeRelease(&pSample);
        return safeBmp;
    }
};

// ==========================================
// 辅助功能：目录扫描等
// ==========================================
bool IsVideoFile(const wstring& path) {
    wstring ext = PathFindExtensionW(path.c_str());
    transform(ext.begin(), ext.end(), ext.begin(), ::towlower);
    return (ext == L".mp4" || ext == L".avi" || ext == L".mov" || ext == L".mkv" || ext == L".wmv" || ext == L".flv" || ext == L".mpg");
}

bool BrowseFolder(HWND hWnd, wstring& outPath) {
    BROWSEINFO bi = { 0 };
    bi.lpszTitle = L"请选择图像保存目录";
    bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE | BIF_USENEWUI;
    bi.hwndOwner = hWnd;
    LPITEMIDLIST pidl = SHBrowseForFolder(&bi);
    if (pidl != 0) {
        WCHAR buffer[MAX_PATH];
        if (SHGetPathFromIDList(pidl, buffer)) { outPath = buffer; }
        CoTaskMemFree(pidl);
        return true;
    }
    return false;
}

void ProcessDrop(const wstring& path) {
    g_batchFiles.clear();

    if (PathIsDirectoryW(path.c_str())) {
        wstring searchPath = path + L"\\*.*";
        WIN32_FIND_DATAW fd;
        HANDLE hFind = FindFirstFileW(searchPath.c_str(), &fd);
        if (hFind != INVALID_HANDLE_VALUE) {
            do {
                if (!(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) {
                    wstring fileName = fd.cFileName;
                    if (IsVideoFile(fileName)) {
                        g_batchFiles.push_back(path + L"\\" + fileName);
                    }
                }
            } while (FindNextFileW(hFind, &fd));
            FindClose(hFind);
        }
    }
    else {
        if (IsVideoFile(path)) g_batchFiles.push_back(path);
    }

    if (g_batchFiles.empty()) {
        MessageBoxW(hMainWnd, L"未找到有效的视频文件！", L"提示", MB_ICONWARNING);
        return;
    }

    g_isConsistent = true;
    g_batchWidth = 0;
    g_batchHeight = 0;

    HCURSOR hOldCursor = SetCursor(LoadCursor(NULL, IDC_WAIT));

    VideoReaderMF tempReader;
    for (size_t i = 0; i < g_batchFiles.size(); ++i) {
        if (FAILED(tempReader.Open(g_batchFiles[i]))) continue;

        UINT32 w, h;
        UINT64 d;
        double f;
        if (SUCCEEDED(tempReader.GetVideoInfo(w, h, d, f))) {
            if (i == 0) {
                g_batchWidth = w;
                g_batchHeight = h;
            }
            else {
                if (w != g_batchWidth || h != g_batchHeight) {
                    g_isConsistent = false;
                }
            }
        }
        tempReader.Close();
    }
    SetCursor(hOldCursor);

    SetDlgItemTextW(hMainWnd, IDC_EDT_PATH, path.c_str());

    if (PathIsDirectoryW(path.c_str())) {
        SetDlgItemTextW(hMainWnd, IDC_EDT_OUT, (path + L"_frames").c_str());
    }
    else {
        WCHAR drive[MAX_PATH], dir[MAX_PATH], name[MAX_PATH], ext[MAX_PATH];
        _wsplitpath_s(path.c_str(), drive, MAX_PATH, dir, MAX_PATH, name, MAX_PATH, ext, MAX_PATH);
        wstring defaultOut = wstring(drive) + wstring(dir) + wstring(name);
        SetDlgItemTextW(hMainWnd, IDC_EDT_OUT, defaultOut.c_str());
    }

    VideoReaderMF reader;
    if (SUCCEEDED(reader.Open(g_batchFiles[0]))) {
        // 初始化全局变量
        g_durationHns = 0;
        g_fps = 0.0;
        g_durationSec = 0.0;
        
        reader.GetVideoInfo(g_batchWidth, g_batchHeight, g_durationHns, g_fps);
        g_durationSec = (double)g_durationHns / 10000000.0;

        reader.Seek(0);
        if (g_pPreviewBitmap) delete g_pPreviewBitmap;
        g_pPreviewBitmap = reader.ReadFrame(g_batchWidth, g_batchHeight);
    }

    BOOL bEnableROI = g_isConsistent;
    EnableWindow(GetDlgItem(hMainWnd, IDC_EDT_X1), bEnableROI);
    EnableWindow(GetDlgItem(hMainWnd, IDC_EDT_Y1), bEnableROI);
    EnableWindow(GetDlgItem(hMainWnd, IDC_EDT_X2), bEnableROI);
    EnableWindow(GetDlgItem(hMainWnd, IDC_EDT_Y2), bEnableROI);

    SetIntToEdit(IDC_EDT_X1, 0);
    SetIntToEdit(IDC_EDT_Y1, 0);
    SetIntToEdit(IDC_EDT_X2, g_batchWidth);
    SetIntToEdit(IDC_EDT_Y2, g_batchHeight);

    // 更新ROI尺寸显示
    UpdateROISizeLabel();

    WCHAR info[256];
    if (g_batchFiles.size() > 1) {
        if (g_isConsistent) {
            swprintf(info, 256, L"批量模式: %zu 个文件 | 分辨率一致 (%dx%d) | 可裁剪", g_batchFiles.size(), g_batchWidth, g_batchHeight);
            SetDlgItemTextW(hMainWnd, IDC_LBL_BATCH, info);
        }
        else {
            swprintf(info, 256, L"批量模式: %zu 个文件 | 分辨率不一致 | 裁剪已禁用", g_batchFiles.size());
            SetDlgItemTextW(hMainWnd, IDC_LBL_BATCH, info);
        }
    }
    else {
        swprintf(info, 256, L"单文件模式 | %dx%d, %.2f秒, %.2f FPS", g_batchWidth, g_batchHeight, g_durationSec, g_fps);
        SetDlgItemTextW(hMainWnd, IDC_LBL_BATCH, info);
    }

    SetWindowSubclass(GetDlgItem(hMainWnd, IDC_PREVIEW), PreviewProc, 0, 0);
    InvalidateRect(GetDlgItem(hMainWnd, IDC_PREVIEW), NULL, FALSE);

    if (SendMessage(GetDlgItem(hMainWnd, IDC_CHK_AUTO), BM_GETCHECK, 0, 0) == BST_CHECKED) {
        StartExtractionThread();
    }
}

// 处理多个拖入的视频文件
void ProcessMultipleFiles(const vector<wstring>& files) {
    g_batchFiles = files;

    if (g_batchFiles.empty()) {
        MessageBoxW(hMainWnd, L"未找到有效的视频文件！", L"提示", MB_ICONWARNING);
        return;
    }

    g_isConsistent = true;
    g_batchWidth = 0;
    g_batchHeight = 0;

    HCURSOR hOldCursor = SetCursor(LoadCursor(NULL, IDC_WAIT));

    // 检查所有文件的分辨率是否一致
    VideoReaderMF tempReader;
    for (size_t i = 0; i < g_batchFiles.size(); ++i) {
        if (FAILED(tempReader.Open(g_batchFiles[i]))) continue;

        UINT32 w, h;
        UINT64 d;
        double f;
        if (SUCCEEDED(tempReader.GetVideoInfo(w, h, d, f))) {
            if (i == 0) {
                g_batchWidth = w;
                g_batchHeight = h;
            }
            else {
                if (w != g_batchWidth || h != g_batchHeight) {
                    g_isConsistent = false;
                }
            }
        }
        tempReader.Close();
    }
    SetCursor(hOldCursor);

    // 获取第一个文件的目录作为输出目录的基础
    wstring firstFile = g_batchFiles[0];
    WCHAR drive[MAX_PATH], dir[MAX_PATH], name[MAX_PATH], ext[MAX_PATH];
    _wsplitpath_s(firstFile.c_str(), drive, MAX_PATH, dir, MAX_PATH, name, MAX_PATH, ext, MAX_PATH);
    wstring baseDir = wstring(drive) + wstring(dir);
    
    // 移除末尾的反斜杠（如果有）
    if (!baseDir.empty() && (baseDir.back() == L'\\' || baseDir.back() == L'/')) {
        baseDir.pop_back();
    }

    // 设置源路径显示为"多文件"提示
    WCHAR pathInfo[128];
    swprintf(pathInfo, 128, L"[多文件模式] 共 %zu 个视频文件", g_batchFiles.size());
    SetDlgItemTextW(hMainWnd, IDC_EDT_PATH, pathInfo);
    
    // 设置输出目录
    SetDlgItemTextW(hMainWnd, IDC_EDT_OUT, (baseDir + L"_frames").c_str());

    // 读取第一个视频的信息用于预览
    VideoReaderMF reader;
    if (SUCCEEDED(reader.Open(g_batchFiles[0]))) {
        g_durationHns = 0;
        g_fps = 0.0;
        g_durationSec = 0.0;
        
        reader.GetVideoInfo(g_batchWidth, g_batchHeight, g_durationHns, g_fps);
        g_durationSec = (double)g_durationHns / 10000000.0;

        reader.Seek(0);
        if (g_pPreviewBitmap) delete g_pPreviewBitmap;
        g_pPreviewBitmap = reader.ReadFrame(g_batchWidth, g_batchHeight);
    }

    // 设置ROI控件状态
    BOOL bEnableROI = g_isConsistent;
    EnableWindow(GetDlgItem(hMainWnd, IDC_EDT_X1), bEnableROI);
    EnableWindow(GetDlgItem(hMainWnd, IDC_EDT_Y1), bEnableROI);
    EnableWindow(GetDlgItem(hMainWnd, IDC_EDT_X2), bEnableROI);
    EnableWindow(GetDlgItem(hMainWnd, IDC_EDT_Y2), bEnableROI);

    SetIntToEdit(IDC_EDT_X1, 0);
    SetIntToEdit(IDC_EDT_Y1, 0);
    SetIntToEdit(IDC_EDT_X2, g_batchWidth);
    SetIntToEdit(IDC_EDT_Y2, g_batchHeight);

    UpdateROISizeLabel();

    // 显示批量模式信息
    WCHAR info[256];
    if (g_isConsistent) {
        swprintf(info, 256, L"批量模式: %zu 个文件 | 分辨率一致 (%dx%d) | 可裁剪", g_batchFiles.size(), g_batchWidth, g_batchHeight);
    }
    else {
        swprintf(info, 256, L"批量模式: %zu 个文件 | 分辨率不一致 | 裁剪已禁用", g_batchFiles.size());
    }
    SetDlgItemTextW(hMainWnd, IDC_LBL_BATCH, info);

    SetWindowSubclass(GetDlgItem(hMainWnd, IDC_PREVIEW), PreviewProc, 0, 0);
    InvalidateRect(GetDlgItem(hMainWnd, IDC_PREVIEW), NULL, FALSE);

    if (SendMessage(GetDlgItem(hMainWnd, IDC_CHK_AUTO), BM_GETCHECK, 0, 0) == BST_CHECKED) {
        StartExtractionThread();
    }
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    hInst = hInstance;

    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    MFStartup(MF_VERSION);

    GdiplusStartupInput gdiplusStartupInput;
    GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);

    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_WIN95_CLASSES | ICC_STANDARD_CLASSES;
    InitCommonControlsEx(&icex);

    WNDCLASSEXW wcex = { 0 };
    wcex.cbSize = sizeof(WNDCLASSEX);
    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = WndProc;
    wcex.hInstance = hInstance;
    wcex.hCursor = LoadCursor(NULL, IDC_ARROW);
    wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wcex.lpszClassName = L"VideoExtractorBatch";
    RegisterClassExW(&wcex);

    hMainWnd = CreateWindowW(L"VideoExtractorBatch", L"drag2frames",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
        CW_USEDEFAULT, 0, 800, 680, NULL, NULL, hInstance, NULL);

    if (!hMainWnd) return FALSE;

    ShowWindow(hMainWnd, nCmdShow);
    UpdateWindow(hMainWnd);

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        if (!IsDialogMessage(hMainWnd, &msg)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    if (g_pPreviewBitmap) delete g_pPreviewBitmap;
    GdiplusShutdown(gdiplusToken);
    MFShutdown();
    CoUninitialize();
    return (int)msg.wParam;
}

int GetIntFromEdit(int id) {
    WCHAR buf[32];
    GetDlgItemTextW(hMainWnd, id, buf, 32);
    return _wtoi(buf);
}

void SetIntToEdit(int id, int val) {
    WCHAR buf[32];
    wsprintfW(buf, L"%d", val);
    SetDlgItemTextW(hMainWnd, id, buf);
}

// 更新ROI尺寸标签
void UpdateROISizeLabel() {
    int x1 = GetIntFromEdit(IDC_EDT_X1);
    int y1 = GetIntFromEdit(IDC_EDT_Y1);
    int x2 = GetIntFromEdit(IDC_EDT_X2);
    int y2 = GetIntFromEdit(IDC_EDT_Y2);
    
    int roiW = x2 - x1;
    int roiH = y2 - y1;
    
    WCHAR buf[64];
    if (roiW > 0 && roiH > 0) {
        swprintf(buf, 64, L"(%dx%d)", roiW, roiH);
    } else {
        wcscpy(buf, L"(无效)");
    }
    SetDlgItemTextW(hMainWnd, IDC_LBL_ROI, buf);
}

void ExtractionWorker(wstring rootOutDir, int interval, RECT roi) {
    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    CLSID jpgClsid;
    GetEncoderClsid(L"image/jpeg", &jpgClsid);

    PostMessage(hMainWnd, WM_USER + 1, 0, 0);
    WCHAR statusBuf[512];

    for (size_t i = 0; i < g_batchFiles.size(); ++i) {
        if (g_stopRequested) break;

        wstring currentFile = g_batchFiles[i];
        swprintf(statusBuf, 512, L"正在处理文件 (%zu/%zu): %s", i + 1, g_batchFiles.size(), PathFindFileNameW(currentFile.c_str()));
        SetDlgItemTextW(hMainWnd, IDC_LBL_INFO, statusBuf);

        WCHAR fName[MAX_PATH], fExt[MAX_PATH];
        _wsplitpath_s(currentFile.c_str(), NULL, 0, NULL, 0, fName, MAX_PATH, fExt, MAX_PATH);
        wstring videoBaseName = fName;  // 保存视频文件名（不含扩展名）
        wstring subOutDir = rootOutDir + L"\\" + fName;

        CreateDirectoryW(subOutDir.c_str(), NULL);

        VideoReaderMF reader;
        if (FAILED(reader.Open(currentFile))) continue;

        UINT32 vW = 0, vH = 0;
        UINT64 vDurHns = 0;
        double vFps = 0.0;
        reader.GetVideoInfo(vW, vH, vDurHns, vFps);
        double vDurSec = (double)vDurHns / 10000000.0;

        int roiW = roi.right - roi.left;
        int roiH = roi.bottom - roi.top;
        if (!g_isConsistent || roiW <= 0 || roiH <= 0) {
            roi.left = 0; roi.top = 0;
            roiW = vW; roiH = vH;
        }

        // 使用顺序读取方式：读取所有帧，按间隔保存
        // interval=0 表示保存每一帧，interval=1 表示每隔1帧保存（即保存第1、3、5...帧）
        int frameIndex = 0;      // 当前读取的帧索引（从0开始）
        int savedCount = 0;      // 已保存的帧数
        int skipCount = interval; // 跳过计数器，初始设为interval以便立即保存第一帧
        
        LONGLONG timestamp = 0;
        
        // 从视频开头开始顺序读取
        reader.Seek(0.0);
        
        while (!g_stopRequested) {
            Bitmap* bmp = reader.ReadNextFrame(vW, vH, &timestamp);
            if (!bmp) break; // 读取完毕或出错
            
            frameIndex++;
            skipCount++;
            
            // 判断是否需要保存这一帧
            if (skipCount > interval) {
                skipCount = 0; // 重置跳过计数器
                
                Bitmap* pSaveBmp = bmp;
                Bitmap* pCropped = nullptr;

                if (g_isConsistent && (roiW < (int)vW || roiH < (int)vH)) {
                    pCropped = bmp->Clone(roi.left, roi.top, roiW, roiH, PixelFormatDontCare);
                    if (pCropped) pSaveBmp = pCropped;
                }

                // 使用"视频文件名_帧序号"格式作为文件名
                WCHAR filePath[MAX_PATH];
                swprintf(filePath, MAX_PATH, L"%s\\%s_%05d.jpg", subOutDir.c_str(), videoBaseName.c_str(), frameIndex);
                pSaveBmp->Save(filePath, &jpgClsid, NULL);

                if (pCropped) delete pCropped;
                savedCount++;
                
                // 更新进度条
                if (savedCount % 5 == 0) {
                    double currentTimeSec = (double)timestamp / 10000000.0;
                    int progress = (int)((currentTimeSec / vDurSec) * 100);
                    if (progress > 100) progress = 100;
                    PostMessage(hMainWnd, WM_USER + 2, progress, 0);
                }
            }
            
            delete bmp;
        }
        reader.Close();
    }

    CoUninitialize();
    PostMessage(hMainWnd, WM_USER + 3, 0, 0);
}

void StartExtractionThread() {
    if (g_batchFiles.empty()) {
        MessageBoxW(hMainWnd, L"没有可处理的文件！", L"错误", MB_ICONERROR);
        return;
    }

    WCHAR outDirBuf[MAX_PATH];
    GetDlgItemTextW(hMainWnd, IDC_EDT_OUT, outDirBuf, MAX_PATH);
    wstring outDir = outDirBuf;
    if (outDir.empty()) {
        MessageBoxW(hMainWnd, L"输出路径不能为空！", L"错误", MB_ICONERROR);
        return;
    }
    CreateDirectoryW(outDir.c_str(), NULL);

    int interval = GetIntFromEdit(IDC_EDT_INT);
    RECT roi;
    roi.left = GetIntFromEdit(IDC_EDT_X1);
    roi.top = GetIntFromEdit(IDC_EDT_Y1);
    roi.right = GetIntFromEdit(IDC_EDT_X2);
    roi.bottom = GetIntFromEdit(IDC_EDT_Y2);

    g_stopRequested = false;
    g_isExtracting = true;
    thread t(ExtractionWorker, outDir, interval, roi);
    t.detach();
}

LRESULT CALLBACK PreviewProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
    if (uMsg == WM_PAINT) {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);
        Graphics graphics(hdc);
        RECT clientRect;
        GetClientRect(hWnd, &clientRect);
        SolidBrush bgBrush(Color(255, 200, 200, 200));
        graphics.FillRectangle(&bgBrush, 0, 0, clientRect.right, clientRect.bottom);
        if (g_pPreviewBitmap && g_batchWidth > 0 && g_batchHeight > 0) {
            float ratioW = (float)clientRect.right / g_batchWidth;
            float ratioH = (float)clientRect.bottom / g_batchHeight;
            // 使用 std::min 计算缩放比例
            float ratio = std::min(ratioW, ratioH);
            int dispW = (int)(g_batchWidth * ratio);
            int dispH = (int)(g_batchHeight * ratio);
            int offX = (clientRect.right - dispW) / 2;
            int offY = (clientRect.bottom - dispH) / 2;
            graphics.DrawImage(g_pPreviewBitmap, offX, offY, dispW, dispH);

            if (g_isConsistent) {
                int x1 = GetIntFromEdit(IDC_EDT_X1);
                int y1 = GetIntFromEdit(IDC_EDT_Y1);
                int x2 = GetIntFromEdit(IDC_EDT_X2);
                int y2 = GetIntFromEdit(IDC_EDT_Y2);
                if (x2 > x1 && y2 > y1) {
                    Pen redPen(Color(255, 255, 0, 0), 2);
                    float drawX = offX + x1 * ratio;
                    float drawY = offY + y1 * ratio;
                    float drawW = (x2 - x1) * ratio;
                    float drawH = (y2 - y1) * ratio;
                    graphics.DrawRectangle(&redPen, drawX, drawY, drawW, drawH);
                }
            }
        }
        EndPaint(hWnd, &ps);
        return 0;
    }
    return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_CREATE:
    {
        int y = 10;
        CreateWindowW(L"STATIC", L"请拖入 [视频文件] 或 [目录] ...", WS_VISIBLE | WS_CHILD, 10, y, 760, 20, hWnd, (HMENU)IDC_LBL_INFO, hInst, NULL);
        y += 30;
        CreateWindowW(L"STATIC", L"源路径:", WS_VISIBLE | WS_CHILD, 10, y, 70, 20, hWnd, NULL, hInst, NULL);
        CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL | ES_READONLY | WS_TABSTOP, 80, y, 670, 20, hWnd, (HMENU)IDC_EDT_PATH, hInst, NULL);
        y += 30;
        CreateWindowW(L"STATIC", L"输出目录:", WS_VISIBLE | WS_CHILD, 10, y, 70, 20, hWnd, NULL, hInst, NULL);
        CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL | WS_TABSTOP, 80, y, 590, 20, hWnd, (HMENU)IDC_EDT_OUT, hInst, NULL);
        CreateWindowW(L"BUTTON", L"...", WS_VISIBLE | WS_CHILD | WS_TABSTOP, 680, y - 2, 70, 24, hWnd, (HMENU)IDC_BTN_BROWSE, hInst, NULL);

        y += 30;
        CreateWindowW(L"STATIC", L"跳帧数:", WS_VISIBLE | WS_CHILD, 10, y, 70, 20, hWnd, NULL, hInst, NULL);
        CreateWindowW(L"EDIT", L"0", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_NUMBER | WS_TABSTOP, 80, y, 50, 20, hWnd, (HMENU)IDC_EDT_INT, hInst, NULL);

        CreateWindowW(L"STATIC", L"ROI区域:", WS_VISIBLE | WS_CHILD, 150, y, 60, 20, hWnd, NULL, hInst, NULL);
        CreateWindowW(L"EDIT", L"0", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_NUMBER | WS_TABSTOP, 210, y, 40, 20, hWnd, (HMENU)IDC_EDT_X1, hInst, NULL);
        CreateWindowW(L"EDIT", L"0", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_NUMBER | WS_TABSTOP, 255, y, 40, 20, hWnd, (HMENU)IDC_EDT_Y1, hInst, NULL);
        CreateWindowW(L"STATIC", L"-", WS_VISIBLE | WS_CHILD, 300, y, 10, 20, hWnd, NULL, hInst, NULL);
        CreateWindowW(L"EDIT", L"0", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_NUMBER | WS_TABSTOP, 315, y, 40, 20, hWnd, (HMENU)IDC_EDT_X2, hInst, NULL);
        CreateWindowW(L"EDIT", L"0", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_NUMBER | WS_TABSTOP, 360, y, 40, 20, hWnd, (HMENU)IDC_EDT_Y2, hInst, NULL);
        
        // ROI尺寸显示标签
        CreateWindowW(L"STATIC", L"", WS_VISIBLE | WS_CHILD, 405, y, 80, 20, hWnd, (HMENU)IDC_LBL_ROI, hInst, NULL);

        CreateWindowW(L"BUTTON", L"自动开始", WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX | WS_TABSTOP, 500, y, 80, 20, hWnd, (HMENU)IDC_CHK_AUTO, hInst, NULL);
        CreateWindowW(L"BUTTON", L"开始提取", WS_VISIBLE | WS_CHILD | WS_TABSTOP, 600, y - 2, 100, 25, hWnd, (HMENU)IDC_BTN_START, hInst, NULL);

        y += 30;
        CreateWindowW(L"STATIC", L"等待拖入...", WS_VISIBLE | WS_CHILD, 10, y, 760, 20, hWnd, (HMENU)IDC_LBL_BATCH, hInst, NULL);

        y += 25;
        CreateWindowW(L"STATIC", L"", WS_VISIBLE | WS_CHILD | SS_ETCHEDFRAME, 10, y, 760, 420, hWnd, (HMENU)IDC_PREVIEW, hInst, NULL);
        y += 430;
        CreateWindowW(PROGRESS_CLASSW, NULL, WS_VISIBLE | WS_CHILD, 10, y, 760, 20, hWnd, (HMENU)IDC_PROGRESS, hInst, NULL);

        DragAcceptFiles(hWnd, TRUE);
    }
    break;

    case WM_DROPFILES:
    {
        HDROP hDrop = (HDROP)wParam;
        UINT fileCount = DragQueryFileW(hDrop, 0xFFFFFFFF, NULL, 0);
        
        if (fileCount == 1) {
            // 单个文件或目录
            WCHAR filePath[MAX_PATH];
            if (DragQueryFileW(hDrop, 0, filePath, MAX_PATH)) {
                ProcessDrop(filePath);
            }
        }
        else if (fileCount > 1) {
            // 多个文件：收集所有视频文件
            vector<wstring> droppedFiles;
            for (UINT i = 0; i < fileCount; ++i) {
                WCHAR filePath[MAX_PATH];
                if (DragQueryFileW(hDrop, i, filePath, MAX_PATH)) {
                    if (!PathIsDirectoryW(filePath) && IsVideoFile(filePath)) {
                        droppedFiles.push_back(filePath);
                    }
                }
            }
            if (!droppedFiles.empty()) {
                ProcessMultipleFiles(droppedFiles);
            }
            else {
                MessageBoxW(hMainWnd, L"未找到有效的视频文件！", L"提示", MB_ICONWARNING);
            }
        }
        DragFinish(hDrop);
    }
    break;

    case WM_COMMAND:
    {
        int id = LOWORD(wParam);
        int code = HIWORD(wParam);

        if (id == IDC_BTN_BROWSE && code == BN_CLICKED) {
            wstring path;
            if (BrowseFolder(hWnd, path)) {
                SetDlgItemTextW(hWnd, IDC_EDT_OUT, path.c_str());
            }
        }

        if (id == IDC_BTN_START && code == BN_CLICKED) {
            if (g_isExtracting) {
                g_stopRequested = true;
                EnableWindow(GetDlgItem(hWnd, IDC_BTN_START), FALSE);
            }
            else {
                StartExtractionThread();
            }
        }

        if ((id == IDC_EDT_X1 || id == IDC_EDT_Y1 || id == IDC_EDT_X2 || id == IDC_EDT_Y2) && code == EN_CHANGE) {
            InvalidateRect(GetDlgItem(hWnd, IDC_PREVIEW), NULL, FALSE);
            UpdateWindow(GetDlgItem(hWnd, IDC_PREVIEW));
            UpdateROISizeLabel();
        }
    }
    break;

    case WM_USER + 1: // Start
        EnableWindow(GetDlgItem(hWnd, IDC_BTN_START), TRUE);
        SetDlgItemTextW(hWnd, IDC_BTN_START, L"停止");
        EnableWindow(GetDlgItem(hWnd, IDC_EDT_PATH), FALSE);
        EnableWindow(GetDlgItem(hWnd, IDC_EDT_OUT), FALSE);
        EnableWindow(GetDlgItem(hWnd, IDC_BTN_BROWSE), FALSE);
        break;

    case WM_USER + 2: // Progress
        SendMessage(GetDlgItem(hWnd, IDC_PROGRESS), PBM_SETPOS, wParam, 0);
        break;

    case WM_USER + 3: // Finish
        g_isExtracting = false;
        SetDlgItemTextW(hWnd, IDC_BTN_START, L"开始提取");
        SetDlgItemTextW(hWnd, IDC_LBL_INFO, L"所有任务已完成！");
        EnableWindow(GetDlgItem(hWnd, IDC_EDT_PATH), TRUE);
        EnableWindow(GetDlgItem(hWnd, IDC_EDT_OUT), TRUE);
        EnableWindow(GetDlgItem(hWnd, IDC_BTN_BROWSE), TRUE);
        SendMessage(GetDlgItem(hWnd, IDC_PROGRESS), PBM_SETPOS, 100, 0);
        MessageBoxW(hWnd, L"所有任务已完成。", L"提示", MB_OK);
        break;

    case WM_DESTROY:
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

int GetEncoderClsid(const WCHAR* format, CLSID* pClsid) {
    UINT  num = 0, size = 0;
    GetImageEncodersSize(&num, &size);
    if (size == 0) return -1;
    ImageCodecInfo* pImageCodecInfo = (ImageCodecInfo*)(malloc(size));
    if (pImageCodecInfo == NULL) return -1;
    GetImageEncoders(num, size, pImageCodecInfo);
    for (UINT j = 0; j < num; ++j) {
        if (wcscmp(pImageCodecInfo[j].MimeType, format) == 0) {
            *pClsid = pImageCodecInfo[j].Clsid;
            free(pImageCodecInfo);
            return j;
        }
    }
    free(pImageCodecInfo);
    return -1;
}
