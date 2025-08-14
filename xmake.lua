add_files'src/**.cpp'
add_includedirs'C:/Program Files (x86)/Windows Kits/10/Include/10.0.26100.0/cppwinrt'
add_rules'mode.release'
add_syslinks('ole32', 'oleaut32', 'runtimeobject', 'shlwapi')
add_vectorexts'all'
on_load(function(target)
	local version = os.iorun'git describe --dirty --match v* --tags':gsub('[(^v)(\n$)]', '')
	target:set('version', version)
end)
set_encodings'utf-8'
set_exceptions'cxx'
set_fpmodels'strict'
set_kind'shared'
set_languages'cxxlatest'
set_pcxxheader'src/pch.hpp'
set_project'ContextMenu-mklink'
set_toolchains'mingw'
set_warnings('everything', 'pedantic')
target'x64'

on_package(function(target)
	os.cp('src/AppxManifest.xml', target:targetdir())
	local version, commit = target:get'version':match'^(%d+%.%d+%.%d+)%-?(%d*)'
	if commit == '' then
		commit = '0'
	end
	io.gsub(target:targetdir() .. '/AppxManifest.xml', "Version='0.0.0.0'", "Version='" .. version .. '.' .. commit .. "'")
end)

local function format()
	return table.join(os.files'src/**', os.files'**/*.json')
end

task'format'
on_run(function()
	os.runv('clang-format -i --Wno-error=unknown', format())
end)
set_menu{
	description = 'Format source code.',
}

task'format:check'
on_run(function()
	os.execv('clang-format -n -Werror --Wno-error=unknown', format())
end)
set_menu{
	description = 'Check code formatting.',
}
