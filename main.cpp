#define WIN32_LEAN_AND_MEAN
#define NOGDI
#define NOMINMAX
#include <windows.h>
#include <wrl.h>
#include "WebView2.h"

using namespace Microsoft::WRL;

#define CHECK_FAILURE(hr) if (FAILED(hr)) { PostQuitMessage(2); return S_OK; }

// 返回值
// 0: 成功
// 1: 未安装 WebView2 Runtime
// 2: 其他错误
int WINAPI wWinMain(
	_In_ HINSTANCE /*hInstance*/,
	_In_opt_ HINSTANCE /*hPrevInstance*/,
	_In_ LPWSTR /*lpCmdLine*/,
	_In_ int /*nCmdShow*/
) {
	HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
	if (FAILED(hr)) {
		return 2;
	}

	ComPtr<ICoreWebView2Controller> webviewController;

	hr = CreateCoreWebView2Environment(Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
		[&webviewController](HRESULT, ICoreWebView2Environment* env) -> HRESULT {
			// 创建 CoreWebView2Controller，父窗口设为 HWND_MESSAGE，因为无需创建窗口
			env->CreateCoreWebView2Controller(HWND_MESSAGE, Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
				[&webviewController, env](HRESULT, ICoreWebView2Controller* controller) -> HRESULT {
					webviewController = controller;

					ComPtr<ICoreWebView2> webview;
					CHECK_FAILURE(controller->get_CoreWebView2(&webview));

					CHECK_FAILURE(webview->add_NavigationCompleted(Callback<ICoreWebView2NavigationCompletedEventHandler>([env](ICoreWebView2* webview, ICoreWebView2NavigationCompletedEventArgs*) {
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
							L"C:\\Users\\Xu\\Desktop\\test.pdf",
							printSettings.Get(),
							Callback<ICoreWebView2PrintToPdfCompletedHandler>([](HRESULT, BOOL isSuccessful) {
								PostQuitMessage(isSuccessful ? 0 : 2);
								return S_OK;
								}).Get()
						));

						return S_OK;
						}).Get(), nullptr));

					// 打开 HTML
					CHECK_FAILURE(webview->Navigate(L"C:\\Users\\Xu\\Desktop\\test.html"));

					return S_OK;
				}
			).Get());

			return S_OK;
		}
	).Get());

	if (FAILED(hr)) {
		if (hr == HRESULT_FROM_WIN32(ERROR_PRODUCT_UNINSTALLED)) {
			return 1;
		} else {
			return 2;
		}
	}

	// 消息循环
	MSG msg;
	while (GetMessage(&msg, NULL, 0, 0)) {
		DispatchMessage(&msg);
	}

	return (int)msg.wParam;
}
