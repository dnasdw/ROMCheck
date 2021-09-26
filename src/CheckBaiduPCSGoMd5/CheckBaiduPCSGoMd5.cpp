#include <sdw.h>

struct SFileMd5
{
	UString LocalFileName;
	UString LocalFileMd5Upper;
	UString RemoteFileName;
	UString RemoteFileMd5Upper;
};

UString trim(const UString& a_sLine)
{
	wstring sTrimmed = UToW(a_sLine);
	wstring::size_type uPos = sTrimmed.find_first_not_of(L"\n\v\f\r \x85\xA0");
	if (uPos == wstring::npos)
	{
		return USTR("");
	}
	sTrimmed.erase(0, uPos);
	uPos = sTrimmed.find_last_not_of(L"\n\v\f\r \x85\xA0");
	if (uPos != wstring::npos)
	{
		sTrimmed.erase(uPos + 1);
	}
	return WToU(sTrimmed);
}

bool empty(const UString& a_sLine)
{
	return trim(a_sLine).empty();
}

int UMain(int argc, UChar* argv[])
{
	if (argc != 2)
	{
		return 1;
	}
	UString sMd5Path = argv[1];
	UString sMd5ErrorTxtPath = sMd5Path + USTR(".error.txt");
	UString sMd5ErrorBatPath = sMd5Path + USTR(".error.bat");
	UString sMd5ErrorTxtName = sMd5ErrorTxtPath;
	UString::size_type uPos = sMd5ErrorTxtPath.find_last_of(USTR("/\\"));
	if (uPos != UString::npos)
	{
		sMd5ErrorTxtName = sMd5ErrorTxtPath.substr(uPos + 1);
	}
	string sErrorBat;
	sErrorBat += "CHCP 65001\r\n";
	sErrorBat += "PUSHD \"%~dp0\"\r\n";
	sErrorBat += Format("BaiduPCS-Go < \"%s\"\r\n", UToU8(sMd5ErrorTxtName).c_str());
	sErrorBat += "POPD\r\n";
	sErrorBat += "PAUSE\r\n";
	string sErrorTxt;
	FILE* fp = UFopen(sMd5Path.c_str(), USTR("rb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	fseek(fp, 0, SEEK_END);
	u32 uFileSize = ftell(fp);
	fseek(fp, 0, SEEK_SET);
	char* pTemp = new char[uFileSize + 1];
	fread(pTemp, 1, uFileSize, fp);
	fclose(fp);
	pTemp[uFileSize] = 0;
	UString sText = U8ToU(pTemp);
	delete[] pTemp;
	vector<UString> vText = SplitOf(sText, USTR("\r\n"));
	for (vector<UString>::iterator it = vText.begin(); it != vText.end(); ++it)
	{
		*it = trim(*it);
	}
	vector<UString>::iterator itText = remove_if(vText.begin(), vText.end(), empty);
	vText.erase(itText, vText.end());
	vector<SFileMd5> vFileMd5;
	vFileMd5.reserve(vText.size() / 17);
	vector<SFileMd5>::iterator itFileMd5 = vFileMd5.end();
	bool bHasBaiduUserId = false;
	bool bLocal = true;
	for (itText = vText.begin(); itText != vText.end(); ++itText)
	{
		UString& sLine = *itText;
		if (!bHasBaiduUserId)
		{
			UString::size_type uPos0 = sLine.find(USTR(" uid: "));
			if (uPos0 == UString::npos)
			{
				continue;
			}
			uPos0 += UCslen(USTR(" uid: "));
			UString::size_type uPos1 = sLine.find(USTR(", "), uPos0);
			if (uPos1 == UString::npos)
			{
				return 1;
			}
			sErrorTxt += "su " + UToU8(sLine.substr(uPos0, uPos1 - uPos0)) + "\r\n";
			bHasBaiduUserId = true;
			continue;
		}
		if (StartWith(sLine, USTR("[1] - [")))
		{
			if (itFileMd5 == vFileMd5.end() || !itFileMd5->LocalFileName.empty())
			{
				vFileMd5.resize(vFileMd5.size() + 1);
				if (vFileMd5.size() > 1)
				{
					itFileMd5 = vFileMd5.end() - 2;
					UPrintf(USTR("%d:\n"), static_cast<n32>(vFileMd5.size() - 2));
					UPrintf(USTR("L   %") PRIUS USTR(" %") PRIUS USTR("\n"), itFileMd5->LocalFileMd5Upper.c_str(), itFileMd5->LocalFileName.c_str());
					UPrintf(USTR("R   %") PRIUS USTR(" %") PRIUS USTR("\n"), itFileMd5->RemoteFileMd5Upper.c_str(), itFileMd5->RemoteFileName.c_str());
					UPrintf(USTR("\n"));
				}
				itFileMd5 = vFileMd5.end() - 1;
			}
			itFileMd5->LocalFileName = sLine.substr(7, sLine.size() - 9);
			bLocal = true;
		}
		else if (StartWith(sLine, USTR("[0] - [")))
		{
			if (itFileMd5 == vFileMd5.end() || !itFileMd5->RemoteFileName.empty())
			{
				vFileMd5.resize(vFileMd5.size() + 1);
				if (vFileMd5.size() > 1)
				{
					itFileMd5 = vFileMd5.end() - 2;
					UPrintf(USTR("%d:\n"), static_cast<n32>(vFileMd5.size() - 2));
					UPrintf(USTR("L   %") PRIUS USTR(" %") PRIUS USTR("\n"), itFileMd5->LocalFileMd5Upper.c_str(), itFileMd5->LocalFileName.c_str());
					UPrintf(USTR("R   %") PRIUS USTR(" %") PRIUS USTR("\n"), itFileMd5->RemoteFileMd5Upper.c_str(), itFileMd5->RemoteFileName.c_str());
					UPrintf(USTR("\n"));
				}
				itFileMd5 = vFileMd5.end() - 1;
			}
			itFileMd5->RemoteFileName = sLine.substr(7, sLine.size() - 23);
			bLocal = false;
		}
		else if (StartWith(sLine, USTR("md5")))
		{
			if (itFileMd5 == vFileMd5.end())
			{
				return 1;
			}
			vector<UString> vLine = SplitOf(sLine, USTR(" "));
			vector<UString>::iterator itLine = remove_if(vLine.begin(), vLine.end(), empty);
			vLine.erase(itLine, vLine.end());
			if (bLocal)
			{
				if (vLine.size() > 1)
				{
					UString sMd5 = vLine[1];
					transform(sMd5.begin(), sMd5.end(), sMd5.begin(), ::toupper);
					itFileMd5->LocalFileMd5Upper = sMd5;
				}
			}
			else
			{
				if (vLine.size() > 2)
				{
					UString sMd5 = vLine[2];
					transform(sMd5.begin(), sMd5.end(), sMd5.begin(), ::toupper);
					itFileMd5->RemoteFileMd5Upper = sMd5;
				}
			}
		}
	}
	UPrintf(USTR("%d:\n"), static_cast<n32>(vFileMd5.size() - 1));
	UPrintf(USTR("L   %") PRIUS USTR(" %") PRIUS USTR("\n"), itFileMd5->LocalFileMd5Upper.c_str(), itFileMd5->LocalFileName.c_str());
	UPrintf(USTR("R   %") PRIUS USTR(" %") PRIUS USTR("\n"), itFileMd5->RemoteFileMd5Upper.c_str(), itFileMd5->RemoteFileName.c_str());
	UPrintf(USTR("\n"));
	bool bError = false;
	for (n32 i = 0; i < static_cast<n32>(vFileMd5.size()); i++)
	{
		SFileMd5& fileMd5 = vFileMd5[i];
		if (fileMd5.RemoteFileMd5Upper != fileMd5.LocalFileMd5Upper)
		{
			bError = true;
			if (!fileMd5.RemoteFileName.empty())
			{
				sErrorTxt += Format("rm \"%s\"\r\n", UToU8(fileMd5.RemoteFileName).c_str());
			}
			if (!fileMd5.LocalFileName.empty() && !fileMd5.RemoteFileName.empty())
			{
				UString sRemoteDirName = USTR("/");
				uPos = fileMd5.RemoteFileName.rfind(USTR("/"));
				if (uPos != UString::npos)
				{
					sRemoteDirName = fileMd5.RemoteFileName.substr(0, uPos);
				}
				sErrorTxt += Format("upload \"%s\" \"%s\"\r\n", UToU8(fileMd5.LocalFileName).c_str(), UToU8(sRemoteDirName).c_str());
			}
			if (!fileMd5.RemoteFileName.empty())
			{
				sErrorTxt += Format("fixmd5 \"%s\"\r\n", UToU8(fileMd5.RemoteFileName).c_str());
			}
			UPrintf(USTR("error ------------------------------\n"));
			UPrintf(USTR("%d:\n"), i);
			UPrintf(USTR("L   %") PRIUS USTR(" %") PRIUS USTR("\n"), fileMd5.LocalFileMd5Upper.c_str(), fileMd5.LocalFileName.c_str());
			UPrintf(USTR("R   %") PRIUS USTR(" %") PRIUS USTR("\n"), fileMd5.RemoteFileMd5Upper.c_str(), fileMd5.RemoteFileName.c_str());
			UPrintf(USTR("\n"));
		}
	}
	if (bError)
	{
		fp = UFopen(sMd5ErrorTxtPath.c_str(), USTR("wb"), false);
		if (fp == nullptr)
		{
			return 1;
		}
		fwrite(sErrorTxt.c_str(), 1, sErrorTxt.size(), fp);
		fclose(fp);
		fp = UFopen(sMd5ErrorBatPath.c_str(), USTR("wb"), false);
		if (fp == nullptr)
		{
			return 1;
		}
		fwrite(sErrorBat.c_str(), 1, sErrorBat.size(), fp);
		fclose(fp);
		return 1;
	}
	return 0;
}
