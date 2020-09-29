#include <sdw.h>

string trim(const string& a_sLine)
{
	string sTrimmed = a_sLine;
	string::size_type uPos = sTrimmed.find_first_not_of("\t\n\v\f\r \x85\xA0");
	if (uPos == string::npos)
	{
		return "";
	}
	sTrimmed.erase(0, uPos);
	uPos = sTrimmed.find_last_not_of("\t\n\v\f\r \x85\xA0");
	if (uPos != string::npos)
	{
		sTrimmed.erase(uPos + 1);
	}
	return sTrimmed;
}

bool empty(const string& a_sLine)
{
	return trim(a_sLine).empty();
}

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
	pTemp[uFileSize] = '\0';
	string sHtml = pTemp;
	delete[] pTemp;
	vector<string> vLine = SplitOf(sHtml, "\r\n");
	for (vector<string>::iterator it = vLine.begin(); it != vLine.end(); ++it)
	{
		*it = trim(*it);
	}
	vector<string>::iterator itLine = remove_if(vLine.begin(), vLine.end(), empty);
	vLine.erase(itLine, vLine.end());
	map<n32, map<n32, string>> mYearIdName;
	n32 nValueBegin = -1;
	regex rKeyYear("<a href=\"#(20\\d\\d)\".*>(20\\d\\d)</a>", regex_constants::ECMAScript);
	for (n32 i = 0; i < static_cast<n32>(vLine.size()); i++)
	{
		const string& sLine = vLine[i];
		smatch match;
		if (regex_search(sLine, match, rKeyYear))
		{
			if (match[1] == match[2])
			{
				n32 nYear = SToN32(match[1]);
				mYearIdName[nYear];
				nValueBegin = i + 1;
			}
		}
	}
	if (nValueBegin < 0)
	{
		return 1;
	}
	n32 nYear = -1;
	regex rValueYear("<a name=\"(20\\d\\d)\">(20\\d\\d)</a>", regex_constants::ECMAScript);
	regex rIdName("<a href=\"https://.*/(\\d+)/?\".*?>(.*)</a>", regex_constants::ECMAScript);
	for (n32 i = nValueBegin; i < static_cast<n32>(vLine.size()); i++)
	{
		const string& sLine = vLine[i];
		if (sLine == "</div>")
		{
			break;
		}
		smatch match;
		if (regex_search(sLine, match, rValueYear))
		{
			if (match[1] == match[2])
			{
				nYear = SToN32(match[1]);
			}
			continue;
		}
		if (nYear < 0)
		{
			continue;
		}
		if (regex_search(sLine, match, rIdName))
		{
			n32 nId = SToN32(match[1]);
			if (!mYearIdName[nYear].insert(make_pair(nId, match[2])).second)
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
		nYear = itYear->first;
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
