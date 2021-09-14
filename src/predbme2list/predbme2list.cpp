#include <sdw.h>

int UMain(int argc, UChar* argv[])
{
	if (argc != 3)
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
	string sHtml = pTemp;
	delete[] pTemp;
	vector<string> vLine = RegexSplitWith(sHtml, "<div class=\"post\" id=\"");
	map<n32, map<n32, string>> mYearIdName;
	regex rIdYearName("<div class=\"post\" id=\"(\\d+)\">.*?<span class=\"p-time\" data=\"\\d+\" title=\"(20\\d\\d).*?<div class=\"p-c p-c-title\">.*?<a class=\"p-title\" href=\".*?\">(.*?)</a>", regex_constants::ECMAScript);
	for (n32 i = 0; i < static_cast<n32>(vLine.size()); i++)
	{
		const string& sLine = vLine[i];
		smatch match;
		if (regex_search(sLine, match, rIdYearName))
		{
			n32 nId = SToN32(match[1]);
			n32 nYear = SToN32(match[2]);
			if (!mYearIdName[nYear].insert(make_pair(nId, match[3])).second)
			{
				return 1;
			}
		}
	}
	fp = UFopen(argv[2], USTR("wb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	for (map<n32, map<n32, string>>::const_iterator itYear = mYearIdName.begin(); itYear != mYearIdName.end(); ++itYear)
	{
		n32 nYear = itYear->first;
		const map<n32, string>& mIdName = itYear->second;
		if (ftell(fp) != 0)
		{
			fprintf(fp, "\r\n\r\n");
		}
		for (map<n32, string>::const_iterator it = mIdName.begin(); it != mIdName.end(); ++it)
		{
			const string& sName = it->second;
			fprintf(fp, "%04d %s\r\n", nYear, sName.c_str());
		}
	}
	fclose(fp);
	return 0;
}
