add_configfiles'src/AppxManifest.xml'
add_files('src/**.cpp', 'src/**.svg')
add_includedirs'C:/Program Files (x86)/Windows Kits/10/Include/10.0.26100.0/cppwinrt'
add_rules('logo', 'mode.release')
add_shflags('-static-libgcc', '-static-libstdc++', '-Wl,-Bstatic', '-lgcc', '-lstdc++', '-lpthread', '-Wl,-Bdynamic')
add_syslinks('ole32', 'oleaut32', 'runtimeobject', 'shlwapi')
add_vectorexts'all'
on_load(function(target)
	local version, commit = os.iorun'git describe --match v* --tags':match'^v(%d+%.%d+%.%d+)%-?(%d*)'
	if commit == '' then
		commit = '0'
	end
	target:set('version', version .. '.' .. commit)
	target:set('configdir', target:targetdir())
end)
set_configvar('Class', 'AB07ABB7-2731-CFF0-A89E-D7B1B7E31E9E')
set_configvar('Display', 'Context Menu mklink')
set_configvar('Logo', 'Logo.png')
set_configvar('Name', 'ContextMenu.mklink')
set_configvar('Publisher', 'qwerty吃小庄')
set_encodings'utf-8'
set_exceptions'cxx'
set_fpmodels'strict'
set_kind'shared'
set_languages'cxxlatest'
set_pcxxheader'src/pch.hpp'
set_project'ContextMenu-mklink'
set_toolchains'mingw'
set_warnings('everything', 'pedantic')

after_clean(function(target)
	os.rm(target:targetdir())
end)

target'x64'
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

rule'logo'
set_extensions'.svg'
on_buildcmd_file(function (target, batchcmds, sourcefile, opt)
	local targetfile = target:targetdir() .. '/' .. path.basename(sourcefile) .. '.png'
	batchcmds:add_depfiles(sourcefile)
	batchcmds:set_depcache(targetfile)
	batchcmds:show_progress(opt.progress, "${color.build.object}converting %s", sourcefile)
	batchcmds:vrunv('cim -h 1000 -w 1000 png', {
		sourcefile,
		targetfile,
	})
end)
