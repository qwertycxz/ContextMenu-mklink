#include "pch.hpp"
using std::filesystem::exists, std::filesystem::is_directory, std::filesystem::path, std::format, std::wstring, std::wstring_view, winrt::check_bool, winrt::check_hresult, winrt::com_ptr, winrt::get_module_lock, winrt::hresult, winrt::hresult_error,
	winrt::implements, winrt::make, winrt::throw_last_error, winrt::Windows::ApplicationModel::Resources::ResourceLoader;

/**
 * [Resource](https://learn.microsoft.com/en-us/uwp/api/windows.applicationmodel.resources.resourceloader) for the current non-UI-thread context.
 *
 * I18n resources are in `/i18n/`. PRI config is in `/src/pri.xml`.
 */
static const auto resource = ResourceLoader::GetForViewIndependentUse();

/**
 * Get a localized string resource.
 * @param key The resource key
 * @return The localized string
 */
[[nodiscard("Pure function")]]
static const wchar_t* LOC(wstring_view key) {
	return resource.GetString(key).c_str();
}

/**
 * A minimal [`IExplorerCommand`](https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/nf-shobjidl_core-iexplorercommand-invoke) implementation with utility functions for linking.
 *
 * To register context menu item in Windows 11, `IExplorerCommand` is the only option.
 */
struct Command : implements<Command, IExplorerCommand> {
	/**
	 * Initialize all member variables as is.
	 * @param directory The directory where the link will be created
	 * @param target The target file or directory to link to
	 * @param icon The icon to use for the link
	 * @param title The title of the link
	 * @param tip The tooltip for the link
	 * @param executable The executable to use for the link
	 * @param extension The extension of the link file
	 */
	Command(path directory, path target, wstring_view icon, wstring_view title, wstring_view tip, wstring_view executable = L"cmd", wstring_view extension = L"") :
		directory(directory),
		target(target),
		icon(icon),
		title(title),
		tip(tip),
		executable(executable),
		extension(extension) {}

	/**
	 * Enumerate sub-commands. Basic commands has no sub-commands.
	 * @param ppEnum Unused output
	 * @return `E_NOTIMPL`
	 */
	HRESULT EnumSubCommands([[maybe_unused]] IEnumExplorerCommand** ppEnum) {
		return E_NOTIMPL;
	}

	/**
	 * Get the UUID of the command. No need to implement.
	 * @param pguidCommandName Unused output
	 * @return `E_NOTIMPL`
	 */
	HRESULT GetCanonicalName([[maybe_unused]] GUID* pguidCommandName) {
		return E_NOTIMPL;
	}

	/**
	 * Get if the command has sub-commands.
	 * @param pFlags Output `ECF_DEFAULT`
	 * @return `S_OK`
	 */
	HRESULT GetFlags(EXPCMDFLAGS* pFlags) {
		*pFlags = ECF_DEFAULT;
		return S_OK;
	}

	/**
	 * Get the icon of the command.
	 * @param psiItemArray Unused input. Don't care the context
	 * @param ppszIcon Output icon path
	 * @return `S_OK` on success, most likely
	 */
	HRESULT GetIcon([[maybe_unused]] IShellItemArray* psiItemArray, LPWSTR* ppszIcon) {
		return SHStrDupW(icon.data(), ppszIcon);
	}

	/**
	 * Get if the command is enabled.
	 * @param psiItemArray Unused input. Don't care the context
	 * @param fOkToBeSlow Unused input. We are not slow
	 * @param pCmdState Always Output `ECS_ENABLED`
	 * @return `S_OK`
	 */
	HRESULT GetState([[maybe_unused]] IShellItemArray* psiItemArray, [[maybe_unused]] BOOL fOkToBeSlow, EXPCMDSTATE* pCmdState) {
		*pCmdState = ECS_ENABLED;
		return S_OK;
	}

	/**
	 * Get the title of the command.
	 * @param psiItemArray Unused input. Don't care the context
	 * @param ppszName Output command title
	 * @return `S_OK` on success, most likely
	 */
	HRESULT GetTitle([[maybe_unused]] IShellItemArray* psiItemArray, LPWSTR* ppszName) {
		return SHStrDupW(title.data(), ppszName);
	}

	/**
	 * Get the tooltip of the command. Seems unused.
	 * @param psiItemArray Unused input. Don't care the context
	 * @param ppszInfotip Output command tooltip
	 * @return `S_OK` on success, most likely
	 */
	HRESULT GetToolTip([[maybe_unused]] IShellItemArray* psiItemArray, LPWSTR* ppszInfotip) {
		return SHStrDupW(tip.data(), ppszInfotip);
	}

protected:
	/**
	 * The directory where the link will be created.
	 */
	const path directory;
	/**
	 * The target file or directory to link to.
	 */
	const path target;

	/**
	 * Create a link.
	 *
	 * Create another process. If current permissions are insufficient, the link will be created with elevated privileges.
	 *
	 * The link is not guaranteed to be created as users could cancel the privilege operation.
	 *
	 * Don't cover the main logic. Both executable and parameter must be specified.
	 * @param link The link to create
	 * @param parameter The command line parameters
	 * @param test Extra permission to test. Should be the return value of [`CreateFile`](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-createfilew)
	 * @return `S_OK` on command execution, most likely
	 */
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

	/**
	 * Get the link path.
	 *
	 * If there's already a file. The link will be named as `stem (2).ext`, and so on.
	 * @return The link path
	 */
	[[nodiscard("Pure function")]]
	const path getLink() const {
		const auto directory_string = directory.wstring();
		path link = format(L"{}/{}{}", directory_string, target.filename().wstring(), extension);
		const auto target_stem = target.stem().wstring();
		const auto target_extension = target.extension().wstring();
		for (auto i = 2; exists(link); i++) [[unlikely]] {
			link = format(L"{}/{} ({}){}{}", directory_string, target_stem, i, target_extension, extension);
		}
		return link;
	}

private:
	/**
	 * The icon resource, for instance `shell32.dll,-249`
	 */
	const wstring_view icon;
	/**
	 * The title of command
	 */
	const wstring_view title;
	/**
	 * The tooltip of command. Seems unused.
	 */
	const wstring_view tip;
	/**
	 * The executable of command. Default is `cmd`.
	 */
	const wstring_view executable;
	/**
	 * The extension of the link file. Default is empty.
	 *
	 * For instance, `.url` for Internet shortcuts, `.lnk` for shell links.
	 */
	const wstring_view extension;

	/**
	 * Assert the permission of the file.
	 *
	 * If the file is `INVALID_HANDLE_VALUE`, throw an error.
	 * @param file The file handle to check, Should be the return value of `CreateFile`
	 */
	static void assertPermission(void* const file) {
		if (file == INVALID_HANDLE_VALUE) throw_last_error();
		CloseHandle(file);
	}
};

/**
 * Create a [symbolic link](https://learn.microsoft.com/en-us/windows/win32/fileio/symbolic-links) with absolute path.
 */
struct AbsoluteSymbolicLink : Command {
	/**
	 * Initialize all member variables as is.
	 * @param directory The directory where the link will be created
	 * @param target The target file or directory to link to
	 */
	AbsoluteSymbolicLink(path directory, path target) : Command(directory, target, L"shell32.dll,-51380", LOC(L"AbsoluteSymbolicLink.GetTitle"), LOC(L"AbsoluteSymbolicLink.GetToolTip")) {}

	/**
	 * Create symbolic link with absolute path.
	 * @param psiItemArray Unused input. The context is given by the constructor
	 * @param pbc Unused input. The context is given by the constructor
	 * @return `S_OK` on command execution, most likely
	 */
	HRESULT Invoke([[maybe_unused]] IShellItemArray* psiItemArray, [[maybe_unused]] IBindCtx* pbc) {
		wstring_view argument;
		if (is_directory(target)) {
			argument = L"/D";
		}

		const auto link = getLink();
		return createLink(link, format(L"/C mklink {} \"{}\" \"{}\"", argument, link.wstring(), target.wstring()));
	}
};

/**
 * Create a [symbolic link](https://learn.microsoft.com/en-us/windows/win32/fileio/symbolic-links) with relative path.
 */
struct RelativeSymbolicLink : Command {
	/**
	 * Initialize all member variables as is.
	 * @param directory The directory where the link will be created
	 * @param target The target file or directory to link to
	 */
	RelativeSymbolicLink(path directory, path target) : Command(directory, target, L"shell32.dll,-16801", LOC(L"RelativeSymbolicLink.GetTitle"), LOC(L"RelativeSymbolicLink.GetToolTip")) {}

	/**
	 * Get if the command is enabled.
	 * @param psiItemArray Unused input. The context is given by the constructor
	 * @param fOkToBeSlow Unused input. We are not slow
	 * @param pCmdState Output `ECS_ENABLED` if `directory` and `target` are in the same volume, otherwise `ECS_DISABLED`
	 * @return `S_OK`
	 */
	HRESULT GetState([[maybe_unused]] IShellItemArray* psiItemArray, [[maybe_unused]] BOOL fOkToBeSlow, EXPCMDSTATE* pCmdState) {
		if (directory.root_path() == target.root_path()) {
			*pCmdState = ECS_ENABLED;
		}
		else {
			*pCmdState = ECS_DISABLED;
		}
		return S_OK;
	}

	/**
	 * Create symbolic link with relative path.
	 * @param psiItemArray Unused input. The context is given by the constructor
	 * @param pbc Unused input. The context is given by the constructor
	 * @return `S_OK` on command execution, most likely
	 */
	HRESULT Invoke([[maybe_unused]] IShellItemArray* psiItemArray, [[maybe_unused]] IBindCtx* pbc) {
		wstring_view argument;
		if (is_directory(target)) {
			argument = L"/D";
		}

		const auto link = getLink();
		return createLink(link, format(L"/C mklink {} \"{}\" \"{}\"", argument, link.wstring(), target.lexically_relative(directory).wstring()));
	}
};

/**
 * Create a [hard link](https://learn.microsoft.com/en-us/windows/win32/fileio/hard-links-and-junctions#hard-links).
 */
struct HardLink : Command {
	/**
	 * Initialize all member variables as is.
	 * @param directory The directory where the link will be created
	 * @param target The target file or directory to link to
	 */
	HardLink(path directory, path target) : Command(directory, target, L"shell32.dll,-1", LOC(L"HardLink.GetTitle"), LOC(L"HardLink.GetToolTip")) {}

	/**
	 * Get if the command is enabled.
	 * @param psiItemArray Unused input. The context is given by the constructor
	 * @param fOkToBeSlow Unused input. We are not slow
	 * @param pCmdState Output `ECS_ENABLED` if `directory` and `target` are in the same volume and `target` is not a directory, otherwise `ECS_DISABLED`
	 * @return `S_OK`
	 */
	HRESULT GetState([[maybe_unused]] IShellItemArray* psiItemArray, [[maybe_unused]] BOOL fOkToBeSlow, EXPCMDSTATE* pCmdState) {
		if (!is_directory(target) && directory.root_path() == target.root_path()) {
			*pCmdState = ECS_ENABLED;
		}
		else {
			*pCmdState = ECS_DISABLED;
		}
		return S_OK;
	}

	/**
	 * Create a hard link.
	 * @param psiItemArray Unused input. The context is given by the constructor
	 * @param pbc Unused input. The context is given by the constructor
	 * @return `S_OK` on command execution, most likely
	 */
	HRESULT Invoke([[maybe_unused]] IShellItemArray* psiItemArray, [[maybe_unused]] IBindCtx* pbc) {
		const auto link = getLink();
		return createLink(link, format(L"/C mklink /H \"{}\" \"{}\"", link.wstring(), target.wstring()), CreateFileW(target.c_str(), FILE_WRITE_DATA, 0, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
	}
};

/**
 * Create a [directory junction](https://learn.microsoft.com/en-us/windows/win32/fileio/hard-links-and-junctions#junctions).
 *
 * This is a legacy type of symbolic link.
 */
struct DirectoryJunction : Command {
	/**
	 * Initialize all member variables as is.
	 * @param directory The directory where the link will be created
	 * @param target The target file or directory to link to
	 */
	DirectoryJunction(path directory, path target) : Command(directory, target, L"shell32.dll,-4", LOC(L"DirectoryJunction.GetTitle"), LOC(L"DirectoryJunction.GetToolTip")) {}

	/**
	 * Get if the command is enabled.
	 * @param psiItemArray Unused input. The context is given by the constructor
	 * @param fOkToBeSlow Unused input. We are not slow
	 * @param pCmdState Output `ECS_ENABLED` if `target` is a directory, otherwise `ECS_DISABLED`
	 * @return `S_OK`
	 */
	HRESULT GetState([[maybe_unused]] IShellItemArray* psiItemArray, [[maybe_unused]] BOOL fOkToBeSlow, EXPCMDSTATE* pCmdState) {
		if (is_directory(target)) {
			*pCmdState = ECS_ENABLED;
		}
		else {
			*pCmdState = ECS_DISABLED;
		}
		return S_OK;
	}

	/**
	 * Create a directory junction.
	 * @param psiItemArray Unused input. The context is given by the constructor
	 * @param pbc Unused input. The context is given by the constructor
	 * @return `S_OK` on command execution, most likely
	 */
	HRESULT Invoke([[maybe_unused]] IShellItemArray* psiItemArray, [[maybe_unused]] IBindCtx* pbc) {
		const auto link = getLink();
		return createLink(link, format(L"/C mklink /J \"{}\" \"{}\"", link.wstring(), target.wstring()));
	}
};

/**
 * Create an [internet shortcut](https://learn.microsoft.com/en-us/windows/win32/lwef/internet-shortcuts).
 *
 * This is a legacy type of shortcut.
 */
struct InternetShortcut : Command {
	/**
	 * Initialize all member variables as is.
	 * @param directory The directory where the link will be created
	 * @param target The target URL to link to
	 */
	InternetShortcut(path directory, path target) : Command(directory, target, L"shell32.dll,-14", LOC(L"InternetShortcut.GetTitle"), LOC(L"InternetShortcut.GetToolTip"), L"powershell", L".url") {}

	/**
	 * Create an internet shortcut.
	 * @param psiItemArray Unused input. The context is given by the constructor
	 * @param pbc Unused input. The context is given by the constructor
	 * @return `S_OK` on command execution, most likely
	 */
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

/**
 * Create a [shell link](https://learn.microsoft.com/en-us/windows/win32/shell/links).
 *
 * A.K.A shortcut.
 */
struct ShellLink : Command {
	/**
	 * Initialize all member variables as is.
	 * @param directory The directory where the link will be created
	 * @param target The target file or directory to link to
	 */
	ShellLink(path directory, path target) : Command(directory, target, L"shell32.dll,-25", LOC(L"ShellLink.GetTitle"), LOC(L"ShellLink.GetToolTip"), L"powershell", L".lnk") {}

	/**
	 * Create a shell link.
	 * @param psiItemArray Unused input. The context is given by the constructor
	 * @param pbc Unused input. The context is given by the constructor
	 * @return `S_OK` on command execution, most likely
	 */
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

/**
 * Enumerate all sub-commands.
 *
 * This implementation is required to provide the sub-commands for the context menu.
 */
struct Enum : implements<Enum, IEnumExplorerCommand> {
	/**
	 * Initialize all member variables as is.
	 * @param directory The directory where the link will be created
	 * @param target The target file or directory to link to
	 * @param command Current command index
	 */
	Enum(path directory, path target, uint32_t command = 0) : directory(directory), target(target), command(command) {}

	/**
	 * Get the clone of the enum.
	 * @param ppenum Output pointer to the new enum
	 * @return `S_OK` on success, most likely
	 */
	HRESULT Clone(IEnumExplorerCommand** ppenum) {
		return make<Enum>(directory, target, command)->QueryInterface(ppenum);
	}

	/**
	 * Get the next command.
	 * @param celt Number of commands to fetch
	 * @param pUICommand Output pointer to the command array
	 * @param pceltFetched Output pointer to the number of commands fetched
	 * @return `S_OK` if the number of commands fetched is equal to `celt`, `S_FALSE` otherwise
	 */
	HRESULT Next(ULONG celt, IExplorerCommand** pUICommand, ULONG* pceltFetched) {
		auto result = S_OK;
		uint8_t fetched = 0;

		while (celt > fetched) [[likely]] {
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

	/**
	 * Reset the enum to the beginning.
	 */
	HRESULT Reset() {
		command = 0;
		return S_OK;
	}

	/**
	 * Skip a number of commands.
	 * @param celt Number of commands to skip
	 * @return `S_OK`
	 */
	HRESULT Skip(ULONG celt) {
		command += celt;
		return S_OK;
	}

private:
	/**
	 * The directory where the link will be created.
	 */
	const path directory;
	/**
	 * The target file or directory to link to.
	 */
	const path target;
	/**
	 * Current command index.
	 */
	uint32_t command = 0;
};

/**
 * The main mklink context menu command implementation.
 *
 * Also implements the [`IObjectWithSite`](https://learn.microsoft.com/en-us/windows/win32/api/ocidl/nn-ocidl-iobjectwithsite) interface to get context (on which folder user right-clicked).
 */
struct Mklink : implements<Mklink, IExplorerCommand, IObjectWithSite> {
	/**
	 * Enumerate all sub-commands.
	 * @param ppEnum Output pointer to the enum
	 * @return `S_OK` on success, most likely
	 */
	HRESULT EnumSubCommands(IEnumExplorerCommand** ppEnum) {
		return make<Enum>(directory, target)->QueryInterface(ppEnum);
	}

	/**
	 * Get the UUID of the command. No need to implement.
	 * @param pguidCommandName Unused output
	 * @return `E_NOTIMPL`
	 */
	HRESULT GetCanonicalName([[maybe_unused]] GUID* pguidCommandName) {
		return E_NOTIMPL;
	}

	/**
	 * Get if the command has sub-commands.
	 * @param pFlags Output `ECF_HASSUBCOMMANDS`
	 * @return `S_OK`
	 */
	HRESULT GetFlags(EXPCMDFLAGS* pFlags) {
		*pFlags = ECF_HASSUBCOMMANDS;
		return S_OK;
	}

	/**
	 * Get the icon of the command.
	 * @param psiItemArray Unused input. Don't care the context
	 * @param ppszIcon Output icon path
	 * @return `S_OK` on success, most likely
	 */
	HRESULT GetIcon([[maybe_unused]] IShellItemArray* psiItemArray, LPWSTR* ppszIcon) {
		return SHStrDupW(L"shell32.dll,-16769", ppszIcon);
	}

	/**
	 * Get the site set lately.
	 * @param riid Requested interface ID
	 * @param ppvSite Output pointer to the interface
	 * @return `S_OK` on success, most likely
	 */
	HRESULT GetSite(REFIID riid, void** ppvSite) {
		return provider->QueryInterface(riid, ppvSite);
	}

	/**
	 * Get if the command is enabled.
	 *
	 * @param psiItemArray Unused input. `explorer.exe` always called `GetState` with null-`psiItemArray`
	 * @param fOkToBeSlow Unused input. We are not slow
	 * @param pCmdState Output `ECS_ENABLED` if user copied single file or directory, `ECS_DISABLED` otherwise
	 * @return `S_OK`
	 */
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

	/**
	 * Get the title of the command.
	 * @param psiItemArray Unused input. Don't care the context
	 * @param ppszName Output command title
	 * @return `S_OK` on success, most likely
	 */
	HRESULT GetTitle([[maybe_unused]] IShellItemArray* psiItemArray, LPWSTR* ppszName) {
		return SHStrDupW(LOC(L"Mklink.GetTitle"), ppszName);
	}

	/**
	 * Get the tooltip of the command. Seems unused.
	 * @param psiItemArray Unused input. Don't care the context
	 * @param ppszInfotip Output command tooltip
	 * @return `S_OK` on success, most likely
	 */
	HRESULT GetToolTip([[maybe_unused]] IShellItemArray* psiItemArray, LPWSTR* ppszInfotip) {
		return SHStrDupW(LOC(L"Mklink.GetToolTip"), ppszInfotip);
	}

	/**
	 * Invoke the command. Not implemented as we use the sub-commands instead of directly invoking the command.
	 * @param psiItemArray Unused input. Don't care the context
	 * @param pbc Unused input. Don't care the context
	 * @return `E_NOTIMPL`
	 */
	HRESULT Invoke([[maybe_unused]] IShellItemArray* psiItemArray, [[maybe_unused]] IBindCtx* pbc) {
		return E_NOTIMPL;
	}

	/**
	 * Set the site for the command. The context is extracted from the shell's view.
	 * @param pUnkSite Input pointer to the site
	 * @return `S_OK`
	 */
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
	/**
	 * Pointer to the service provider (site).
	 */
	com_ptr<IServiceProvider> provider = nullptr;
	/**
	 * The directory where the link will be created.
	 */
	path directory = path();
	/**
	 * The target file or directory to link to.
	 */
	path target = path();
};

/**
 * The class factory for the COM object. Required.
 */
struct Factory : implements<Factory, IClassFactory> {
	/**
	 * Create an instance of the Mklink class.
	 * @param pUnkOuter Unused input. Don't care the UUID
	 * @param riid Requested interface ID
	 * @param ppvObject Output pointer to the created object
	 * @return `S_OK` on success, most likely
	 */
	HRESULT CreateInstance([[maybe_unused]] IUnknown* pUnkOuter, REFIID riid, void** ppvObject) {
		return make<Mklink>()->QueryInterface(riid, ppvObject);
	}

	/**
	 * Lock or unlock the server.
	 * @param fLock `TRUE` to lock the server, `FALSE` to unlock
	 * @return `S_OK`
	 */
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

/**
 * Check if the DLL can be unloaded.
 * @return `S_OK` if the DLL can be unloaded, `S_FALSE` otherwise
 */
STDAPI DllCanUnloadNow() {
	if (get_module_lock()) return S_FALSE;
	return S_OK;
}

/**
 * The factory of factory (huh?). Required.
 *
 * @param rclsid Unused input. Don't care the UUID
 * @param riid Requested interface ID
 * @param ppv Output pointer to the created object
 * @return `S_OK` on success, most likely
 */
STDAPI DllGetClassObject([[maybe_unused]] REFCLSID rclsid, REFIID riid, LPVOID* ppv) {
	return make<Factory>()->QueryInterface(riid, ppv);
}
