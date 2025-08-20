#include "pch.hpp"
using std::filesystem::exists, std::filesystem::is_directory, std::filesystem::path, std::format, std::wstring, std::wstring_view, winrt::check_bool, winrt::check_hresult, winrt::com_ptr, winrt::get_module_lock, winrt::hresult, winrt::hresult_error,
	winrt::implements, winrt::make, winrt::throw_last_error, winrt::Windows::ApplicationModel::Resources::ResourceLoader;

static const auto resource = ResourceLoader::GetForViewIndependentUse();

[[nodiscard("Pure function")]]
static const wchar_t* LOC(wstring_view key) {
	return resource.GetString(key).c_str();
}

struct Command : implements<Command, IExplorerCommand> {
	Command(path directory, path target, wstring_view icon, wstring_view title, wstring_view tip, wstring_view executable = L"cmd", wstring_view extension = L"") :
		directory(directory),
		target(target),
		icon(icon),
		title(title),
		tip(tip),
		executable(executable),
		extension(extension) {}

	HRESULT EnumSubCommands([[maybe_unused]] IEnumExplorerCommand** ppEnum) {
		return E_NOTIMPL;
	}

	HRESULT GetCanonicalName([[maybe_unused]] GUID* pguidCommandName) {
		return E_NOTIMPL;
	}

	HRESULT GetFlags(EXPCMDFLAGS* pFlags) {
		*pFlags = ECF_DEFAULT;
		return S_OK;
	}

	HRESULT GetIcon([[maybe_unused]] IShellItemArray* psiItemArray, LPWSTR* ppszIcon) {
		return SHStrDupW(icon.data(), ppszIcon);
	}

	HRESULT GetState([[maybe_unused]] IShellItemArray* psiItemArray, [[maybe_unused]] BOOL fOkToBeSlow, EXPCMDSTATE* pCmdState) {
		*pCmdState = ECS_ENABLED;
		return S_OK;
	}

	HRESULT GetTitle([[maybe_unused]] IShellItemArray* psiItemArray, LPWSTR* ppszName) {
		return SHStrDupW(title.data(), ppszName);
	}

	HRESULT GetToolTip([[maybe_unused]] IShellItemArray* psiItemArray, LPWSTR* ppszInfotip) {
		return SHStrDupW(tip.data(), ppszInfotip);
	}

protected:
	const path directory;
	const path target;

	[[nodiscard("Please handle error")]]
	const hresult createLink(const path link, const wstring_view parameter, void* const test = nullptr) const {
		wstring_view operation;
		try {
			assertPermission(test);
			assertPermission(CreateFileW(link.c_str(), FILE_WRITE_DATA, 0, nullptr, CREATE_NEW, FILE_ATTRIBUTE_TEMPORARY | FILE_FLAG_DELETE_ON_CLOSE, nullptr));
		}
		catch (const hresult_error e) {
			const auto code = e.code();
			if (uint16_t(code) != ERROR_ACCESS_DENIED) [[unlikely]] {
				MessageBoxW(nullptr, e.message().c_str(), LOC(L"Command.Error"), MB_ICONERROR);
				return code;
			}
			operation = L"runas";
		}

		ShellExecuteW(nullptr, operation.data(), executable.data(), parameter.data(), nullptr, SW_HIDE);
		return S_OK;
	}

	[[nodiscard("Pure function")]]
	const path getLink() const {
		const auto directory_string = directory.wstring();
		path link = format(L"{}/{}{}", directory_string, target.filename().wstring(), extension);
		const auto target_stem = target.stem().wstring();
		const auto target_extension = target.extension().wstring();
		for (auto i = 2; exists(link); i++) {
			link = format(L"{}/{} ({}){}{}", directory_string, target_stem, i, target_extension, extension);
		}
		return link;
	}

private:
	const wstring_view icon;
	const wstring_view title;
	const wstring_view tip;
	const wstring_view executable;
	const wstring_view extension;

	static void assertPermission(void* const file) {
		if (file == INVALID_HANDLE_VALUE) throw_last_error();
		CloseHandle(file);
	}
};

struct AbsoluteSymbolicLink : Command {
	AbsoluteSymbolicLink(path directory, path target) : Command(directory, target, L"shell32.dll,-51380", LOC(L"AbsoluteSymbolicLink.GetTitle"), LOC(L"AbsoluteSymbolicLink.GetToolTip")) {}

	HRESULT Invoke([[maybe_unused]] IShellItemArray* psiItemArray, [[maybe_unused]] IBindCtx* pbc) {
		wstring_view argument;
		if (is_directory(target)) {
			argument = L"/D";
		}

		const auto link = getLink();
		return createLink(link, format(L"/C mklink {} \"{}\" \"{}\"", argument, link.wstring(), target.wstring()));
	}
};

struct RelativeSymbolicLink : Command {
	RelativeSymbolicLink(path directory, path target) : Command(directory, target, L"shell32.dll,-16801", LOC(L"RelativeSymbolicLink.GetTitle"), LOC(L"RelativeSymbolicLink.GetToolTip")) {}

	HRESULT GetState([[maybe_unused]] IShellItemArray* psiItemArray, [[maybe_unused]] BOOL fOkToBeSlow, EXPCMDSTATE* pCmdState) {
		if (directory.root_path() == target.root_path()) {
			*pCmdState = ECS_ENABLED;
		}
		else {
			*pCmdState = ECS_DISABLED;
		}
		return S_OK;
	}

	HRESULT Invoke([[maybe_unused]] IShellItemArray* psiItemArray, [[maybe_unused]] IBindCtx* pbc) {
		wstring_view argument;
		if (is_directory(target)) {
			argument = L"/D";
		}

		const auto link = getLink();
		return createLink(link, format(L"/C mklink {} \"{}\" \"{}\"", argument, link.wstring(), target.lexically_relative(directory).wstring()));
	}
};

struct HardLink : Command {
	HardLink(path directory, path target) : Command(directory, target, L"shell32.dll,-1", LOC(L"HardLink.GetTitle"), LOC(L"HardLink.GetToolTip")) {}

	HRESULT GetState([[maybe_unused]] IShellItemArray* psiItemArray, [[maybe_unused]] BOOL fOkToBeSlow, EXPCMDSTATE* pCmdState) {
		if (!is_directory(target) && directory.root_path() == target.root_path()) {
			*pCmdState = ECS_ENABLED;
		}
		else {
			*pCmdState = ECS_DISABLED;
		}
		return S_OK;
	}

	HRESULT Invoke([[maybe_unused]] IShellItemArray* psiItemArray, [[maybe_unused]] IBindCtx* pbc) {
		const auto link = getLink();
		return createLink(link, format(L"/C mklink /H \"{}\" \"{}\"", link.wstring(), target.wstring()), CreateFileW(target.c_str(), FILE_WRITE_DATA, 0, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
	}
};

struct DirectoryJunction : Command {
	DirectoryJunction(path directory, path target) : Command(directory, target, L"shell32.dll,-4", LOC(L"DirectoryJunction.GetTitle"), LOC(L"DirectoryJunction.GetToolTip")) {}

	HRESULT GetState([[maybe_unused]] IShellItemArray* psiItemArray, [[maybe_unused]] BOOL fOkToBeSlow, EXPCMDSTATE* pCmdState) {
		if (is_directory(target)) {
			*pCmdState = ECS_ENABLED;
		}
		else {
			*pCmdState = ECS_DISABLED;
		}
		return S_OK;
	}

	HRESULT Invoke([[maybe_unused]] IShellItemArray* psiItemArray, [[maybe_unused]] IBindCtx* pbc) {
		const auto link = getLink();
		return createLink(link, format(L"/C mklink /J \"{}\" \"{}\"", link.wstring(), target.wstring()));
	}
};

struct InternetShortcut : Command {
	InternetShortcut(path directory, path target) : Command(directory, target, L"shell32.dll,-14", LOC(L"InternetShortcut.GetTitle"), LOC(L"InternetShortcut.GetToolTip"), L"powershell", L".url") {}

	HRESULT Invoke([[maybe_unused]] IShellItemArray* psiItemArray, [[maybe_unused]] IBindCtx* pbc) {
		const auto link = getLink();
		// clang-format off
		return createLink(link, format(L"-Command New-Item '{}' -Value '\
			[InternetShortcut]                                        \n\
			URL = {}                                                  \n\
		'", link.wstring(), target.wstring()));
		// clang-format on
	}
};

struct ShellLink : Command {
	ShellLink(path directory, path target) : Command(directory, target, L"shell32.dll,-25", LOC(L"ShellLink.GetTitle"), LOC(L"ShellLink.GetToolTip"), L"powershell", L".lnk") {}

	HRESULT Invoke([[maybe_unused]] IShellItemArray* psiItemArray, [[maybe_unused]] IBindCtx* pbc) {
		const auto link = getLink();
		// clang-format off
		return createLink(link, format(L"-Command                                  \
			$shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut('{}');\
			$shortcut.TargetPath = '{}';                                           \
			$shortcut.Save();                                                      \
		", link.wstring(), target.wstring()));
		// clang-format on
	}
};

struct Enum : implements<Enum, IEnumExplorerCommand> {
	Enum(path directory, path target) : directory(directory), target(target) {}

	Enum(path directory, path target, uint32_t command) : directory(directory), target(target), command(command) {}

	HRESULT Clone(IEnumExplorerCommand** ppenum) {
		return make<Enum>(directory, target, command)->QueryInterface(ppenum);
	}

	HRESULT Next(ULONG celt, IExplorerCommand** pUICommand, ULONG* pceltFetched) {
		auto result = S_OK;
		uint8_t fetched = 0;

		while (celt > fetched) {
			switch (command + fetched) {
				case 0:
					result = make<AbsoluteSymbolicLink>(directory, target)->QueryInterface(fetched + pUICommand);
					break;
				case 1:
					result = make<RelativeSymbolicLink>(directory, target)->QueryInterface(fetched + pUICommand);
					break;
				case 2:
					result = make<HardLink>(directory, target)->QueryInterface(fetched + pUICommand);
					break;
				case 3:
					result = make<DirectoryJunction>(directory, target)->QueryInterface(fetched + pUICommand);
					break;
				case 4:
					result = make<InternetShortcut>(directory, target)->QueryInterface(fetched + pUICommand);
					break;
				case 5:
					result = make<ShellLink>(directory, target)->QueryInterface(fetched + pUICommand);
					break;
				default:
					result = S_FALSE;
			}

			if (result != S_OK) break;
			fetched++;
		}
		command += fetched;

		if (pceltFetched != nullptr) [[unlikely]] {
			*pceltFetched = fetched;
		}
		return result;
	}

	HRESULT Reset() {
		command = 0;
		return S_OK;
	}

	HRESULT Skip(ULONG celt) {
		command += celt;
		return S_OK;
	}

private:
	const path directory;
	const path target;
	uint32_t command = 0;
};

struct Mklink : implements<Mklink, IExplorerCommand, IObjectWithSite> {
	HRESULT EnumSubCommands(IEnumExplorerCommand** ppEnum) {
		return make<Enum>(directory, target)->QueryInterface(ppEnum);
	}

	HRESULT GetCanonicalName([[maybe_unused]] GUID* pguidCommandName) {
		return E_NOTIMPL;
	}

	HRESULT GetFlags(EXPCMDFLAGS* pFlags) {
		*pFlags = ECF_HASSUBCOMMANDS;
		return S_OK;
	}

	HRESULT GetIcon([[maybe_unused]] IShellItemArray* psiItemArray, LPWSTR* ppszIcon) {
		return SHStrDupW(L"shell32.dll,-16769", ppszIcon);
	}

	HRESULT GetSite(REFIID riid, void** ppvSite) {
		return provider->QueryInterface(riid, ppvSite);
	}

	HRESULT GetState([[maybe_unused]] IShellItemArray* psiItemArray, [[maybe_unused]] BOOL fOkToBeSlow, EXPCMDSTATE* pCmdState) {
		if (directory.empty() || !OpenClipboard(nullptr)) [[unlikely]] {
			*pCmdState = ECS_DISABLED;
			return S_OK;
		}

		auto data = HDROP(GetClipboardData(CF_HDROP));
		if (DragQueryFileW(data, 0xFFFFFFFF, nullptr, 0) == 1) [[unlikely]] {
			*pCmdState = ECS_ENABLED;
			wstring target_string(32767, 0);
			DragQueryFileW(data, 0, target_string.data(), 32767);
			target = target_string.c_str();
		}
		else {
			*pCmdState = ECS_DISABLED;
		}

		CloseClipboard();
		return S_OK;
	}

	HRESULT GetTitle([[maybe_unused]] IShellItemArray* psiItemArray, LPWSTR* ppszName) {
		return SHStrDupW(LOC(L"Mklink.GetTitle"), ppszName);
	}

	HRESULT GetToolTip([[maybe_unused]] IShellItemArray* psiItemArray, LPWSTR* ppszInfotip) {
		return SHStrDupW(LOC(L"Mklink.GetToolTip"), ppszInfotip);
	}

	HRESULT Invoke([[maybe_unused]] IShellItemArray* psiItemArray, [[maybe_unused]] IBindCtx* pbc) {
		return E_NOTIMPL;
	}

	HRESULT SetSite(IUnknown* pUnkSite) {
		if (pUnkSite == nullptr) [[unlikely]] {
			directory.clear();
			provider = nullptr;
			return S_OK;
		}

		ITEMIDLIST* list = nullptr;
		try {
			check_hresult(pUnkSite->QueryInterface(provider.put()));
			com_ptr<IShellBrowser> browser;
			check_hresult(provider->QueryService(IID_IShellBrowser, browser.put()));
			com_ptr<IShellView> shell;
			check_hresult(browser->QueryActiveShellView(shell.put()));
			com_ptr<IPersistFolder2> folder;
			check_hresult(shell.as<IFolderView>()->GetFolder(IID_IPersistFolder2, folder.put_void()));
			check_hresult(folder->GetCurFolder(&list));
			wstring link_string(32767, 0);
			check_bool(SHGetPathFromIDListW(list, link_string.data()));
			directory = link_string.c_str();
		}
		catch (...) {
			directory.clear();
		}

		CoTaskMemFree(list);
		return S_OK;
	}

private:
	com_ptr<IServiceProvider> provider = nullptr;
	path directory = path();
	path target = path();
};

struct Factory : implements<Factory, IClassFactory> {
	HRESULT CreateInstance([[maybe_unused]] IUnknown* pUnkOuter, REFIID riid, void** ppvObject) {
		return make<Mklink>()->QueryInterface(riid, ppvObject);
	}

	HRESULT LockServer(BOOL fLock) {
		if (fLock) {
			++get_module_lock();
		}
		else {
			--get_module_lock();
		}
		return S_OK;
	}
};

STDAPI DllCanUnloadNow() {
	if (get_module_lock()) return S_FALSE;
	return S_OK;
}

STDAPI DllGetClassObject([[maybe_unused]] REFCLSID rclsid, REFIID riid, LPVOID* ppv) {
	return make<Factory>()->QueryInterface(riid, ppv);
}
