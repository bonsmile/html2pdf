#define WIN32_LEAN_AND_MEAN
#define NOGDI
#define NOMINMAX
#include <windows.h>
#include <wrl.h>
#include "WebView2.h"
#include "shellapi.h"
#include <string>

using namespace Microsoft::WRL;

#define CHECK_FAILURE(hr) if (FAILED(hr)) { PostQuitMessage(1); return S_OK; }

static bool FileExists(const wchar_t* fileName) noexcept {
	DWORD attrs = GetFileAttributes(fileName);
	// 排除文件夹
	return (attrs != INVALID_FILE_ATTRIBUTES) && !(attrs & FILE_ATTRIBUTE_DIRECTORY);
}

// 返回值
// 0: 成功
// 1: 其他错误
// 2: 未安装 WebView2 Runtime
// 3: 解析命令行失败
// 4: 文件不存在
int WINAPI wWinMain(
	_In_ HINSTANCE /*hInstance*/,
	_In_opt_ HINSTANCE /*hPrevInstance*/,
	_In_ LPWSTR lpCmdLine,
	_In_ int /*nCmdShow*/
) {
	int nArgs;
	wchar_t** argList = CommandLineToArgvW(lpCmdLine, &nArgs);
	if (!argList || nArgs < 1 || nArgs > 2) {
		return 3;
	}

	std::wstring src = argList[0];
	if (!FileExists(src.c_str())) {
		return 4;
	}

	std::wstring dest;
	if (nArgs == 2) {
		dest = argList[1];
	} else {
		size_t pointPos = src.find_last_of(L'.');
		if (pointPos == std::wstring::npos) {
			dest = src + L".pdf";
		} else {
			dest = src.substr(0, pointPos + 1) + L"pdf";
		}
	}
	
	LocalFree(argList);

	HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
	if (FAILED(hr)) {
		return 1;
	}

	ComPtr<ICoreWebView2Controller> webviewController;

	hr = CreateCoreWebView2Environment(Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
		[&webviewController, &src, &dest](HRESULT, ICoreWebView2Environment* env) -> HRESULT {
			// 创建 CoreWebView2Controller，父窗口设为 HWND_MESSAGE，因为无需创建窗口
			env->CreateCoreWebView2Controller(HWND_MESSAGE, Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
				[&webviewController, env, &src, &dest](HRESULT, ICoreWebView2Controller* controller) -> HRESULT {
					webviewController = controller;

					ComPtr<ICoreWebView2> webview;
					CHECK_FAILURE(controller->get_CoreWebView2(&webview));

					CHECK_FAILURE(webview->add_NavigationCompleted(Callback<ICoreWebView2NavigationCompletedEventHandler>([env, &dest](ICoreWebView2* webview, ICoreWebView2NavigationCompletedEventArgs*) {
						ComPtr<ICoreWebView2_7> webview2_7;
						CHECK_FAILURE(webview->QueryInterface<ICoreWebView2_7>(webview2_7.GetAddressOf()));

						ComPtr<ICoreWebView2PrintSettings> printSettings;
						{
							ComPtr<ICoreWebView2Environment6> webviewEnvironment6;
							CHECK_FAILURE(env->QueryInterface(
								IID_PPV_ARGS(&webviewEnvironment6)));
							CHECK_FAILURE(webviewEnvironment6->CreatePrintSettings(&printSettings));
						}

						printSettings->put_ShouldPrintBackgrounds(TRUE);
						// A4 尺寸，单位为英尺
						printSettings->put_PageWidth(8.268);
						printSettings->put_PageHeight(11.693);

						// 导出 PDF
						CHECK_FAILURE(webview2_7->PrintToPdf(
							dest.c_str(),
							printSettings.Get(),
							Callback<ICoreWebView2PrintToPdfCompletedHandler>([](HRESULT, BOOL isSuccessful) {
								PostQuitMessage(isSuccessful ? 0 : 2);
								return S_OK;
							}).Get()
						));

						return S_OK;
					}).Get(), nullptr));

					// 打开 HTML
					CHECK_FAILURE(webview->Navigate(src.c_str()));

					return S_OK;
				}
			).Get());

			return S_OK;
		}
	).Get());

	if (FAILED(hr)) {
		return hr == HRESULT_FROM_WIN32(ERROR_PRODUCT_UNINSTALLED) ? 2 : 1;
	}

	// 消息循环
	MSG msg;
	while (GetMessage(&msg, NULL, 0, 0)) {
		DispatchMessage(&msg);
	}

	return (int)msg.wParam;
}
