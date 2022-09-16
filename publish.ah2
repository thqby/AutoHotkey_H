github := GitHubRelease('thqby/AutoHotkey_H', FileRead('E:\Documents\github access token.txt'))
tagname := 'v' ComObject('Shell.Application').NameSpace(A_ScriptDir '\bin\x64w_MT').ParseName('AutoHotkey.exe').ExtendedProperty('ProductVersion')
release_folder := 'bin\release'

if (!A_Args.Length) {
	pack()

	r := github.create(tagname)
	if !(r := github.upload_asset(tagname, release_folder '\' tagname '.7z', tagname '.7z')).Has('id')
		throw JSON.stringify(r)
} else {
	switch A_Args[1] {
		case 'del':	; del tag asset?
			if A_Args.Has(3)
				github.delete_asset(A_Args[3], A_Args[2])
			else github.delete(A_Args[2])
		case 'upload':	; upload tag filename name label?
			if !(r := github.upload_asset(A_Args[2], A_Args[3], A_Args[4], A_Args.Has(5) ? A_Args[5] : unset)).Has('id')
				throw JSON.stringify(r)
		case 'pack':
			pack()
	}
}

pack() {
	SetWorkingDir(A_ScriptDir)
	DirCreate(release_folder)
	FileDelete(release_folder '\*.*')
	try FileCopy('bin\Win32w_MT\AutoHotkey.exe', release_folder '\AutoHotkey32.exe', true)
	try FileCopy('bin\Win32w_MT\AutoHotkey.dll', release_folder '\AutoHotkey32.dll', true)
	try FileCopy('bin\x64w_MT\AutoHotkey.exe', release_folder '\AutoHotkey64.exe', true)
	try FileCopy('bin\x64w_MT\AutoHotkey.dll', release_folder '\AutoHotkey64.dll', true)
	try FileCopy('source\ahkapi.h', release_folder '\ahkapi.h', true)
	RunWait('"C:\Program Files\7-Zip\7z.exe" a -mx9 ' tagname '.7z autohotkey*', release_folder, 'hide')
}

class GitHubRelease {
	__New(owner_repo, access_token := '') {
		this.api := 'https://api.github.com/repos/' owner_repo '/releases'
		this.access_token := access_token
		this.req := ComObject('WinHttp.WinHttpRequest.5.1')
	}

	request(url, verb := 'GET', data?, accept := 'application/vnd.github+json') {
		(req := this.req).Open(verb, A_Clipboard := this.api url)
		req.SetRequestHeader('Accept', accept)
		if this.access_token
			req.SetRequestHeader('Authorization', 'token ' this.access_token)
		req.Send(data?)
		if InStr(accept, 'json')
			try return JSON.parse(req.ResponseText)
		return req.ResponseText
	}

	create(tag_name, name?, target_commitish?, body?, draft := false, prerelease := false) {
		return this.request('', 'POST', JSON.stringify({
			tag_name: tag_name,
			name: name ?? tag_name,
			target_commitish: target_commitish?,
			body: body?,
			draft: draft ? JSON.true : JSON.false,
			prerelease: prerelease ? JSON.true : JSON.false
		}))
	}

	get(tag_or_id?) {
		if !IsSet(tag_or_id)
			return this.request('/latest')
		if tag_or_id is Integer
			return this.request(tag_or_id ? '/' tag_or_id : '')
		return this.request('/tags/' tag_or_id)
	}

	update(tag_or_id, tag_name?, name?, target_commitish?, body?, draft?, prerelease?) {
		return this.request('/' this.get_tagid(tag_or_id), 'PATCH', JSON.stringify({
			tag_name: tag_name?,
			name: name?,
			target_commitish: target_commitish?,
			body: body?,
			draft: IsSet(draft) ? (draft ? JSON.true : JSON.false) : unset,
			prerelease: IsSet(prerelease) ? (prerelease ? JSON.true : JSON.false) : unset
		}))
	}

	delete(tag_or_id) {
		this.request('/' this.get_tagid(tag_or_id), 'DELETE')
	}

	get_tagid(tag) {
		if tag is String
			tag := this.request('/tags/' tag).Get('id', 0)
		if !tag
			throw TargetError()
		return tag
	}

	upload_asset(tag_or_id, filepath, name, label?) {
		if !this.access_token
			throw Error('miss access token')
		tag_or_id := this.get_tagid(tag_or_id)
		try this.delete_asset(name, tag_or_id)
		url := 'https://uploads' SubStr(this.api, 12) '/' tag_or_id '/assets?name=' name (IsSet(label) ? '&label=' label : '')
		; curlcmd := 'cmd /c curl -X POST -H "Accept: application/vnd.github+json" -H "Authorization: token ' this.access_token '" -F "file=@' filepath '" "' url '" > out'
		; RunWait(curlcmd, A_ScriptDir, 'hide'), ret := FileRead(f := A_ScriptDir '\out'), FileDelete(f)
		; return JSON.parse(ret)
		(req := this.req).Open('POST', url)
		req.SetRequestHeader('Accept', 'application/vnd.github+json')
		req.SetRequestHeader('Authorization', 'token ' this.access_token)
		data := ReadFileToStream(filepath)
		pvData := NumGet(ComObjValue(data) + 8 + A_PtrSize, 'ptr')
		cbElements := data.MaxIndex() + 1
		DllCall('urlmon\FindMimeFromData', 'ptr', 0, 'ptr', 0, 'ptr', pvData, 'uint', Min(256, cbElements), 'ptr', 0, 'uint', 0, 'str*', &mime, 'uint', 0, 'HRESULT')
		req.SetRequestHeader('Content-Type', mime)
		req.SetRequestHeader('Content-Length', String(cbElements))
		req.Send(data)
		return JSON.parse(req.ResponseText)
	}

	get_assetid(name, tag_or_id?) {
		if name is Integer
			return name
		for asset in this.request('/' this.get_tagid(tag_or_id) '/assets') {
			if asset['name'] = name
				return asset['id']
		}
		throw TargetError()
	}

	delete_asset(name_or_id, tag_or_id?) {
		return this.request('/assets/' this.get_assetid(name_or_id, tag_or_id?), 'DELETE')
	}
}

ReadFileToStream(path) {
	ado := ComObject('Adodb.Stream')
	ado.Type := 1
	ado.Open()
	ado.LoadFromFile(path)
	stream := ado.Read(), ado.Close()
	return stream
}