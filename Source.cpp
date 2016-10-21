#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#pragma comment(lib, "oleaut32")
#pragma comment(lib,"gdiplus")
#pragma comment(lib,"shlwapi")

#include <windows.h>
#include <shlwapi.h>
#include <gdiplus.h>
#include <olectl.h>

TCHAR szClassName[] = TEXT("Window");

typedef struct
{
	WORD idReserved;
	WORD idType;
	WORD idCount;
} ICONHEADER;

typedef struct
{
	BYTE bWidth;
	BYTE bHeight;
	BYTE bColorCount;
	BYTE bReserved;
	WORD wPlanes;
	WORD wBitCount;
	DWORD dwBytesInRes;
	DWORD dwImageOffset;
} ICONDIR;

typedef struct
{
	BITMAPINFOHEADER biHeader;
} ICONIMAGE;

static UINT WriteIconHeader(HANDLE hFile, int nImages)
{
	ICONHEADER iconheader;
	DWORD nWritten;
	iconheader.idReserved = 0;
	iconheader.idType = 1;
	iconheader.idCount = nImages;
	WriteFile(hFile, &iconheader, sizeof(iconheader), &nWritten, 0);
	return nWritten;
}

static UINT NumBitmapBytes(BITMAP *pBitmap)
{
	int nWidthBytes = pBitmap->bmWidthBytes;
	if (nWidthBytes & 3)
		nWidthBytes = (nWidthBytes + 4) & ~3;
	return nWidthBytes * pBitmap->bmHeight;
}

static UINT WriteIconImageHeader(HANDLE hFile, BITMAP *pbmpColor, BITMAP *pbmpMask)
{
	BITMAPINFOHEADER biHeader = { 0 };
	DWORD nWritten;
	UINT nImageBytes;
	nImageBytes = NumBitmapBytes(pbmpColor) + NumBitmapBytes(pbmpMask);
	ZeroMemory(&biHeader, sizeof(biHeader));
	biHeader.biSize = sizeof(biHeader);
	biHeader.biWidth = pbmpColor->bmWidth;
	biHeader.biHeight = pbmpColor->bmHeight * 2;
	biHeader.biPlanes = pbmpColor->bmPlanes;
	biHeader.biBitCount = pbmpColor->bmBitsPixel;
	biHeader.biSizeImage = nImageBytes;
	WriteFile(hFile, &biHeader, sizeof(biHeader), &nWritten, 0);
	return nWritten;
}

static BOOL GetIconBitmapInfo(HICON hIcon, ICONINFO *pIconInfo, BITMAP *pbmpColor, BITMAP *pbmpMask)
{
	if (!GetIconInfo(hIcon, pIconInfo))
		return FALSE;
	if (!GetObject(pIconInfo->hbmColor, sizeof(BITMAP), pbmpColor))
		return FALSE;
	if (!GetObject(pIconInfo->hbmMask, sizeof(BITMAP), pbmpMask))
		return FALSE;
	return TRUE;
}

static UINT WriteIconDirectoryEntry(HANDLE hFile, int nIdx, HICON hIcon, UINT nImageOffset)
{
	ICONINFO iconInfo = { 0 };
	ICONDIR iconDir = { 0 };
	BITMAP bmpColor = { 0 };
	BITMAP bmpMask = { 0 };
	DWORD nWritten;
	UINT nColorCount;
	UINT nImageBytes;
	GetIconBitmapInfo(hIcon, &iconInfo, &bmpColor, &bmpMask);
	nImageBytes = NumBitmapBytes(&bmpColor) + NumBitmapBytes(&bmpMask);
	if (bmpColor.bmBitsPixel >= 8)
		nColorCount = 0;
	else
		nColorCount = 1 << (bmpColor.bmBitsPixel * bmpColor.bmPlanes);
	iconDir.bWidth = (BYTE)bmpColor.bmWidth;
	iconDir.bHeight = (BYTE)bmpColor.bmHeight;
	iconDir.bColorCount = nColorCount;
	iconDir.bReserved = 0;
	iconDir.wPlanes = bmpColor.bmPlanes;
	iconDir.wBitCount = bmpColor.bmBitsPixel;
	iconDir.dwBytesInRes = sizeof(BITMAPINFOHEADER) + nImageBytes;
	iconDir.dwImageOffset = nImageOffset;
	WriteFile(hFile, &iconDir, sizeof(iconDir), &nWritten, 0);
	DeleteObject(iconInfo.hbmColor);
	DeleteObject(iconInfo.hbmMask);
	return nWritten;
}

static UINT WriteIconData(HANDLE hFile, HBITMAP hBitmap)
{
	BITMAP bmp = { 0 };
	int i;
	BYTE * pIconData;
	UINT nBitmapBytes;
	DWORD nWritten;
	GetObject(hBitmap, sizeof(BITMAP), &bmp);
	nBitmapBytes = NumBitmapBytes(&bmp);
	pIconData = (BYTE *)malloc(nBitmapBytes);
	GetBitmapBits(hBitmap, nBitmapBytes, pIconData);
	for (i = bmp.bmHeight - 1; i >= 0; i--)
	{
		WriteFile(
			hFile,
			pIconData + (i * bmp.bmWidthBytes),
			bmp.bmWidthBytes,
			&nWritten,
			0);
		if (bmp.bmWidthBytes & 3)
		{
			DWORD padding = 0;
			WriteFile(hFile, &padding, 4 - (bmp.bmWidthBytes & 3), &nWritten, 0);
		}
	}
	free(pIconData);
	return nBitmapBytes;
}

BOOL SaveIcon3(TCHAR *szIconFile, HICON hIcon[], int nNumIcons)
{
	int * pImageOffset;
	if (hIcon == 0 || nNumIcons < 1)
		return FALSE;
	HANDLE hFile = CreateFile(szIconFile, GENERIC_WRITE, 0, 0, CREATE_ALWAYS, 0, 0);
	if (hFile == INVALID_HANDLE_VALUE)
		return FALSE;
	WriteIconHeader(hFile, nNumIcons);
	SetFilePointer(hFile, sizeof(ICONDIR) * nNumIcons, 0, FILE_CURRENT);
	pImageOffset = (int *)malloc(nNumIcons * sizeof(int));
	for (int i = 0; i < nNumIcons; ++i)
	{
		ICONINFO iconInfo = { 0 };
		BITMAP bmpColor = { 0 }, bmpMask = { 0 };
		GetIconBitmapInfo(hIcon[i], &iconInfo, &bmpColor, &bmpMask);
		pImageOffset[i] = SetFilePointer(hFile, 0, 0, FILE_CURRENT);
		WriteIconImageHeader(hFile, &bmpColor, &bmpMask);
		WriteIconData(hFile, iconInfo.hbmColor);
		WriteIconData(hFile, iconInfo.hbmMask);
		DeleteObject(iconInfo.hbmColor);
		DeleteObject(iconInfo.hbmMask);
	}
	SetFilePointer(hFile, sizeof(ICONHEADER), 0, FILE_BEGIN);
	for (int i = 0; i < nNumIcons; ++i)
	{
		WriteIconDirectoryEntry(hFile, i, hIcon[i], pImageOffset[i]);
	}
	free(pImageOffset);
	CloseHandle(hFile);
	return TRUE;
}

HICON CreateAlphaIcon(Gdiplus::Bitmap *pImg, DWORD dwSize)
{
	BITMAPV5HEADER bi = { 0 };
	void *lpBits;
	HICON hAlphaIcon = NULL;
	ZeroMemory(&bi, sizeof(BITMAPV5HEADER));
	bi.bV5Size = sizeof(BITMAPV5HEADER);
	bi.bV5Width = dwSize;
	bi.bV5Height = dwSize;
	bi.bV5Planes = 1;
	bi.bV5BitCount = 32;
	bi.bV5Compression = BI_BITFIELDS;
	bi.bV5RedMask = 0x00FF0000;
	bi.bV5GreenMask = 0x0000FF00;
	bi.bV5BlueMask = 0x000000FF;
	bi.bV5AlphaMask = 0xFF000000;
	bi.bV5CSType = LCS_WINDOWS_COLOR_SPACE;
	bi.bV5Endpoints.ciexyzRed.ciexyzX =
		bi.bV5Endpoints.ciexyzGreen.ciexyzX =
		bi.bV5Endpoints.ciexyzBlue.ciexyzX = 0;
	bi.bV5Endpoints.ciexyzRed.ciexyzY =
		bi.bV5Endpoints.ciexyzGreen.ciexyzY =
		bi.bV5Endpoints.ciexyzBlue.ciexyzY = 0;
	bi.bV5Endpoints.ciexyzRed.ciexyzZ =
		bi.bV5Endpoints.ciexyzGreen.ciexyzZ =
		bi.bV5Endpoints.ciexyzBlue.ciexyzZ = 0;
	bi.bV5GammaRed = 0;
	bi.bV5GammaGreen = 0;
	bi.bV5GammaBlue = 0;
	bi.bV5Intent = LCS_GM_IMAGES;
	bi.bV5ProfileData = 0;
	bi.bV5ProfileSize = 0;
	bi.bV5Reserved = 0;
	const HDC hdc = GetDC(NULL);
	const HBITMAP hBitmap = CreateDIBSection(hdc, reinterpret_cast<BITMAPINFO *>(&bi),
		DIB_RGB_COLORS,
		&lpBits, NULL, DWORD(0));
	const HDC hMemDC = CreateCompatibleDC(hdc);
	ReleaseDC(NULL, hdc);
	const HBITMAP hOldBitmap = static_cast<HBITMAP>(SelectObject(hMemDC, hBitmap));
	{
		Gdiplus::Bitmap *bitmap = new Gdiplus::Bitmap(dwSize, dwSize, PixelFormat32bppARGB);
		{
			Gdiplus::Graphics * g = Gdiplus::Graphics::FromImage(bitmap);
			g->SetInterpolationMode(Gdiplus::InterpolationModeHighQuality);
			g->SetCompositingQuality(Gdiplus::CompositingQualityHighQuality);
			g->SetCompositingMode(Gdiplus::CompositingModeSourceOver);
			g->DrawImage(pImg, Gdiplus::Rect(0, 0, dwSize, dwSize),
				0, 0,
				pImg->GetWidth(), pImg->GetHeight(),
				Gdiplus::UnitPixel);
			delete g;
		}
		Gdiplus::BitmapData bitmapdata;
		const Gdiplus::Rect rect(0, 0, dwSize, dwSize);
		bitmap->LockBits(&rect, Gdiplus::ImageLockModeRead, PixelFormat32bppARGB, &bitmapdata);
		const int stride = bitmapdata.Stride;
		const UINT* pixel = (UINT*)bitmapdata.Scan0;
		const byte* p = (byte*)pixel;
		{
			DWORD *lpdwPixel;
			lpdwPixel = (DWORD *)lpBits;
			for (int y = dwSize - 1; y >=0; --y)
			{
				for (int x = 0; x < (int)dwSize; ++x)
				{
					for (int i = 0; i < 4; ++i)
					{
						*lpdwPixel |= (p[4 * x + i + y*stride] << (8*i));
					}
					lpdwPixel++;
				}
			}
		}
		bitmap->UnlockBits(&bitmapdata);
		delete bitmap;
	}
	SelectObject(hMemDC, hOldBitmap);
	DeleteDC(hMemDC);
	HBITMAP hMonoBitmap = ::CreateBitmap(dwSize, dwSize, 1, 1, NULL);
	ICONINFO ii = { 0 };
	ii.fIcon = TRUE;
	ii.xHotspot = 0;
	ii.yHotspot = 0;
	ii.hbmMask = hMonoBitmap;
	ii.hbmColor = hBitmap;
	hAlphaIcon = ::CreateIconIndirect(&ii);
	DeleteObject(hBitmap);
	DeleteObject(hMonoBitmap);
	return hAlphaIcon;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	switch (msg)
	{
	case WM_CREATE:
		CreateWindow(TEXT("BUTTON"), TEXT("16 x 16 (&1)"), WS_VISIBLE | WS_CHILD | WS_TABSTOP | BS_AUTOCHECKBOX, 10, 10, 256, 32, hWnd, (HMENU)100, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		CreateWindow(TEXT("BUTTON"), TEXT("32 x 32 (&2)"), WS_VISIBLE | WS_CHILD | WS_TABSTOP | BS_AUTOCHECKBOX, 10, 50, 256, 32, hWnd, (HMENU)101, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		CreateWindow(TEXT("BUTTON"), TEXT("48 x 48 (&3)"), WS_VISIBLE | WS_CHILD | WS_TABSTOP | BS_AUTOCHECKBOX, 10, 90, 256, 32, hWnd, (HMENU)102, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		CreateWindow(TEXT("BUTTON"), TEXT("64 x 64 (&4)"), WS_VISIBLE | WS_CHILD | WS_TABSTOP | BS_AUTOCHECKBOX, 10, 130, 256, 32, hWnd, (HMENU)103, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		CreateWindow(TEXT("BUTTON"), TEXT("96 x 96 (&5)"), WS_VISIBLE | WS_CHILD | WS_TABSTOP | BS_AUTOCHECKBOX, 10, 170, 256, 32, hWnd, (HMENU)104, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		CreateWindow(TEXT("BUTTON"), TEXT("128 x 128 (&6)"), WS_VISIBLE | WS_CHILD | WS_TABSTOP | BS_AUTOCHECKBOX, 10, 210, 256, 32, hWnd, (HMENU)105, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		CreateWindow(TEXT("BUTTON"), TEXT("256 x 256 (&7)"), WS_VISIBLE | WS_CHILD | WS_TABSTOP | BS_AUTOCHECKBOX, 10, 250, 256, 32, hWnd, (HMENU)106, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		CreateWindow(TEXT("STATIC"), TEXT("チェックを付けて画像ファイルをドラッグ＆ドロップしてください。"), WS_VISIBLE | WS_CHILD, 10, 290, 512, 32, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		DragAcceptFiles(hWnd, TRUE);
		break;
	case WM_DROPFILES:
	{
		HDROP hDrop = (HDROP)wParam;
		TCHAR szFileName[MAX_PATH];
		UINT iFile, nFiles;
		nFiles = DragQueryFile((HDROP)hDrop, 0xFFFFFFFF, NULL, 0);
		for (iFile = 0; iFile<nFiles; ++iFile)
		{
			DragQueryFile(hDrop, iFile, szFileName, sizeof(szFileName));
			Gdiplus::Bitmap *bitmap = Gdiplus::Bitmap::FromFile(szFileName);
			if (bitmap)
			{
				HICON hIcon[64] = { 0 };
				int nTotalCount = 0;
				int listSize[] = { 16, 32, 48, 64, 96, 128, 256 };
				for (int i = 0; i <= 6; ++i)
				{
					if (SendDlgItemMessage(hWnd, 100 + i, BM_GETCHECK, 0, 0))
					{
						hIcon[nTotalCount] = CreateAlphaIcon(bitmap, listSize[i]);
						++nTotalCount;
					}
				}
				TCHAR szOutputIconPath[MAX_PATH];
				lstrcpy(szOutputIconPath, szFileName);
				PathRemoveExtension(szOutputIconPath);
				PathAddExtension(szOutputIconPath, TEXT(".ico"));
				SaveIcon3(szOutputIconPath, hIcon, nTotalCount);
				for (int i = 0; i < nTotalCount; ++i)
				{
					DestroyIcon(hIcon[i]);
				}
				delete bitmap;
			}
		}
		DragFinish(hDrop);
	}
	break;
	case WM_CLOSE:
		DestroyWindow(hWnd);
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return(DefDlgProc(hWnd, msg, wParam, lParam));
	}
	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPSTR pCmdLine, int nCmdShow)
{
	ULONG_PTR gdiToken;
	Gdiplus::GdiplusStartupInput gdiSI;
	Gdiplus::GdiplusStartup(&gdiToken, &gdiSI, NULL);
	MSG msg;
	WNDCLASS wndclass = {
		CS_HREDRAW | CS_VREDRAW,
		WndProc,
		0,
		DLGWINDOWEXTRA,
		hInstance,
		0,
		LoadCursor(0, IDC_ARROW),
		(HBRUSH)(COLOR_WINDOW + 1),
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("画像データからアイコンを作成する"),
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
	);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	while (GetMessage(&msg, 0, 0, 0))
	{
		if (!IsDialogMessage(hWnd, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	Gdiplus::GdiplusShutdown(gdiToken);
	return (int)msg.wParam;
}
