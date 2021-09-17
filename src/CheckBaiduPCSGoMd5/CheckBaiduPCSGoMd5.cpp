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
	FILE* fp = UFopen(argv[1], USTR("rb"), false);
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
	bool bLocal = true;
	for (itText = vText.begin(); itText != vText.end(); ++itText)
	{
		UString& sLine = *itText;
		if (StartWith(sLine, USTR("[1] - [")))
		{
			if (itFileMd5 == vFileMd5.end() || !itFileMd5->LocalFileName.empty())
			{
				vFileMd5.resize(vFileMd5.size() + 1);
				if (vFileMd5.size() > 1)
				{
					itFileMd5 = vFileMd5.end() - 2;
					UPrintf(USTR("%d:\n"), vFileMd5.size() - 2);
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
					UPrintf(USTR("%d:\n"), vFileMd5.size() - 2);
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
	UPrintf(USTR("%d:\n"), vFileMd5.size() - 1);
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
			UPrintf(USTR("error ------------------------------\n"));
			UPrintf(USTR("%d:\n"), i);
			UPrintf(USTR("L   %") PRIUS USTR(" %") PRIUS USTR("\n"), fileMd5.LocalFileMd5Upper.c_str(), fileMd5.LocalFileName.c_str());
			UPrintf(USTR("R   %") PRIUS USTR(" %") PRIUS USTR("\n"), fileMd5.RemoteFileMd5Upper.c_str(), fileMd5.RemoteFileName.c_str());
			UPrintf(USTR("\n"));
		}
	}
	return bError ? 1 : 0;
}
