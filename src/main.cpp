#include "pch.hpp"
using std::filesystem::exists, std::filesystem::is_directory, std::filesystem::path, std::format, std::wstring, std::wstring_view, winrt::check_bool, winrt::check_hresult, winrt::com_ptr, winrt::get_module_lock, winrt::hresult, winrt::hresult_error,
	winrt::implements, winrt::make, winrt::throw_last_error;

struct Command : implements<Command, IExplorerCommand> {
	Command(path directory, path target, wstring_view icon, wstring_view title, wstring_view tip, wstring_view executable = L"cmd", wstring_view extension = L"") :
		directory(directory),
		target(target),
		icon(icon),
		title(title),
		tip(tip),
		executable(executable),
		extension(extension) {}

	HRESULT EnumSubCommands(IEnumExplorerCommand** ppEnum) {
		(void)ppEnum;
		return E_NOTIMPL;
	}

	HRESULT GetCanonicalName(GUID* pguidCommandName) {
		(void)pguidCommandName;
		return E_NOTIMPL;
	}

	HRESULT GetFlags(EXPCMDFLAGS* pFlags) {
		*pFlags = ECF_DEFAULT;
		return S_OK;
	}

	HRESULT GetIcon(IShellItemArray* psiItemArray, LPWSTR* ppszIcon) {
		(void)psiItemArray;
		return SHStrDupW(icon.data(), ppszIcon);
	}

	HRESULT GetState(IShellItemArray* psiItemArray, BOOL fOkToBeSlow, EXPCMDSTATE* pCmdState) {
		(void)psiItemArray;
		(void)fOkToBeSlow;

		*pCmdState = ECS_ENABLED;
		return S_OK;
	}

	HRESULT GetTitle(IShellItemArray* psiItemArray, LPWSTR* ppszName) {
		(void)psiItemArray;
		return SHStrDupW(title.data(), ppszName);
	}

	HRESULT GetToolTip(IShellItemArray* psiItemArray, LPWSTR* ppszInfotip) {
		(void)psiItemArray;
		return SHStrDupW(tip.data(), ppszInfotip);
	}

protected:
	const path directory;
	const path target;

	const hresult createLink(const path link, const wstring_view parameter, void* test = nullptr) const {
		wstring_view operation;
		try {
			assertPermission(test);
			assertPermission(CreateFileW(link.c_str(), FILE_WRITE_DATA, 0, nullptr, CREATE_NEW, FILE_ATTRIBUTE_TEMPORARY | FILE_FLAG_DELETE_ON_CLOSE, nullptr));
		}
		catch (const hresult_error e) {
			const auto code = e.code();
			if (uint16_t(code) != ERROR_ACCESS_DENIED) {
				MessageBoxW(nullptr, e.message().c_str(), L"Error", MB_ICONERROR);
				return code;
			}
			operation = L"runas";
		}

		ShellExecuteW(nullptr, operation.data(), executable.data(), parameter.data(), nullptr, SW_HIDE);
		return S_OK;
	}

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

	void assertPermission(void* file) const {
		if (file == INVALID_HANDLE_VALUE) throw_last_error();
		CloseHandle(file);
	}
};

struct AbsoluteSymbolicLink : Command {
	AbsoluteSymbolicLink(path directory, path target) : Command(directory, target, L"shell32.dll,-51380", L"Symbolic link (Absolute)", L"No restrictions") {}

	HRESULT Invoke(IShellItemArray* psiItemArray, IBindCtx* pbc) {
		(void)psiItemArray;
		(void)pbc;

		wstring_view argument;
		if (is_directory(target)) {
			argument = L"/D";
		}

		const auto link = getLink();
		return createLink(link, format(L"/C mklink {} \"{}\" \"{}\"", argument, link.wstring(), target.wstring()));
	}
};

struct RelativeSymbolicLink : Command {
	RelativeSymbolicLink(path directory, path target) : Command(directory, target, L"shell32.dll,-16801", L"Symbolic link (Relative)", L"Same volume only") {}

	HRESULT GetState(IShellItemArray* psiItemArray, BOOL fOkToBeSlow, EXPCMDSTATE* pCmdState) {
		(void)psiItemArray;
		(void)fOkToBeSlow;

		if (directory.root_path() == target.root_path()) {
			*pCmdState = ECS_ENABLED;
		}
		else {
			*pCmdState = ECS_DISABLED;
		}
		return S_OK;
	}

	HRESULT Invoke(IShellItemArray* psiItemArray, IBindCtx* pbc) {
		(void)psiItemArray;
		(void)pbc;

		wstring_view argument;
		if (is_directory(target)) {
			argument = L"/D";
		}

		const auto link = getLink();
		return createLink(link, format(L"/C mklink {} \"{}\" \"{}\"", argument, link.wstring(), target.lexically_relative(directory).wstring()));
	}
};

struct HardLink : Command {
	HardLink(path directory, path target) : Command(directory, target, L"shell32.dll,-1", L"Hard link", L"Files and same volume only") {}

	HRESULT GetState(IShellItemArray* psiItemArray, BOOL fOkToBeSlow, EXPCMDSTATE* pCmdState) {
		(void)psiItemArray;
		(void)fOkToBeSlow;

		if (!is_directory(target) && directory.root_path() == target.root_path()) {
			*pCmdState = ECS_ENABLED;
		}
		else {
			*pCmdState = ECS_DISABLED;
		}
		return S_OK;
	}

	HRESULT Invoke(IShellItemArray* psiItemArray, IBindCtx* pbc) {
		(void)psiItemArray;
		(void)pbc;

		const auto link = getLink();
		return createLink(link, format(L"/C mklink /H \"{}\" \"{}\"", link.wstring(), target.wstring()), CreateFileW(target.c_str(), FILE_WRITE_DATA, 0, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
	}
};

struct DirectoryJunction : Command {
	DirectoryJunction(path directory, path target) : Command(directory, target, L"shell32.dll,-4", L"Directory Junction", L"Directories and same computer only") {}

	HRESULT GetState(IShellItemArray* psiItemArray, BOOL fOkToBeSlow, EXPCMDSTATE* pCmdState) {
		(void)psiItemArray;
		(void)fOkToBeSlow;

		if (is_directory(target)) {
			*pCmdState = ECS_ENABLED;
		}
		else {
			*pCmdState = ECS_DISABLED;
		}
		return S_OK;
	}

	HRESULT Invoke(IShellItemArray* psiItemArray, IBindCtx* pbc) {
		(void)psiItemArray;
		(void)pbc;

		const auto link = getLink();
		return createLink(link, format(L"/C mklink /J \"{}\" \"{}\"", link.wstring(), target.wstring()));
	}
};

struct InternetShortcut : Command {
	InternetShortcut(path directory, path target) : Command(directory, target, L"shell32.dll,-14", L"Shortcut (.url)", L"Internet Shortcut", L"powershell", L".url") {}

	HRESULT Invoke(IShellItemArray* psiItemArray, IBindCtx* pbc) {
		(void)psiItemArray;
		(void)pbc;

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
	ShellLink(path directory, path target) : Command(directory, target, L"shell32.dll,-25", L"Shortcut (.lnk)", L"Shell link", L"powershell", L".lnk") {}

	HRESULT Invoke(IShellItemArray* psiItemArray, IBindCtx* pbc) {
		(void)psiItemArray;
		(void)pbc;

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

		if (pceltFetched != nullptr) {
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

	HRESULT GetCanonicalName(GUID* pguidCommandName) {
		(void)pguidCommandName;
		return E_NOTIMPL;
	}

	HRESULT GetFlags(EXPCMDFLAGS* pFlags) {
		*pFlags = ECF_HASSUBCOMMANDS;
		return S_OK;
	}

	HRESULT GetIcon(IShellItemArray* psiItemArray, LPWSTR* ppszIcon) {
		(void)psiItemArray;
		return SHStrDupW(L"shell32.dll,-16769", ppszIcon);
	}

	HRESULT GetSite(REFIID riid, void** ppvSite) {
		return provider->QueryInterface(riid, ppvSite);
	}

	HRESULT GetState(IShellItemArray* psiItemArray, BOOL fOkToBeSlow, EXPCMDSTATE* pCmdState) {
		(void)psiItemArray;
		(void)fOkToBeSlow;

		if (directory.empty() || !OpenClipboard(nullptr)) {
			*pCmdState = ECS_DISABLED;
			return S_OK;
		}

		auto data = HDROP(GetClipboardData(CF_HDROP));
		if (DragQueryFileW(data, 0xFFFFFFFF, nullptr, 0) == 1) {
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

	HRESULT GetTitle(IShellItemArray* psiItemArray, LPWSTR* ppszName) {
		(void)psiItemArray;
		return SHStrDupW(L"Create link", ppszName);
	}

	HRESULT GetToolTip(IShellItemArray* psiItemArray, LPWSTR* ppszInfotip) {
		(void)psiItemArray;
		return SHStrDupW(L"Create a link to copied directory or file", ppszInfotip);
	}

	HRESULT Invoke(IShellItemArray* psiItemArray, IBindCtx* pbc) {
		(void)psiItemArray;
		(void)pbc;
		return E_NOTIMPL;
	}

	HRESULT SetSite(IUnknown* pUnkSite) {
		if (pUnkSite == nullptr) {
			directory.clear();
			provider = nullptr;
			return S_OK;
		}

		PIDLIST_ABSOLUTE list = nullptr;
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
	HRESULT CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppvObject) {
		(void)pUnkOuter;
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

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv) {
	(void)rclsid;
	return make<Factory>()->QueryInterface(riid, ppv);
}
