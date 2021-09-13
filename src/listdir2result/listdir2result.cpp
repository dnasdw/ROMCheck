#include <sdw.h>

struct SRecord
{
	n32 Year;
	string Name;
	string Path;
	bool Exist;
};

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
	if (argc != 5)
	{
		return 1;
	}
	n32 nDepthMax = SToN32(argv[3]);
	map<string, UString> mDir;
	queue<pair<UString, n32>> qDir;
	qDir.push(make_pair(argv[2], 0));
	while (!qDir.empty())
	{
		UString& sParent = qDir.front().first;
		n32 nDepth = qDir.front().second;
		if (nDepth > nDepthMax)
		{
			qDir.pop();
			continue;
		}
#if SDW_PLATFORM == SDW_PLATFORM_WINDOWS
		WIN32_FIND_DATAW ffd;
		HANDLE hFind = INVALID_HANDLE_VALUE;
		wstring sPattern = sParent + L"/*";
		hFind = FindFirstFileW(sPattern.c_str(), &ffd);
		if (hFind != INVALID_HANDLE_VALUE)
		{
			do
			{
				if ((ffd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0 && wcscmp(ffd.cFileName, L".") != 0 && wcscmp(ffd.cFileName, L"..") != 0)
				{
					wstring sDir = sParent + L"/" + ffd.cFileName;
					if (nDepth == nDepthMax)
					{
						if (!mDir.insert(make_pair(UToU8(ffd.cFileName), sDir)).second)
						{
							return 1;
						}
					}
					if (nDepth + 1 <= nDepthMax)
					{
						qDir.push(make_pair(sDir, nDepth + 1));
					}
				}
			} while (FindNextFileW(hFind, &ffd) != 0);
			FindClose(hFind);
		}
#else
		DIR* pDir = opendir(sParent.c_str());
		if (pDir != nullptr)
		{
			dirent* pDirent = nullptr;
			while ((pDirent = readdir(pDir)) != nullptr)
			{
				string sName = pDirent->d_name;
#if SDW_PLATFORM == SDW_PLATFORM_MACOS
				sName = TSToS<string, string>(sName, "UTF-8-MAC", "UTF-8");
#endif
				// handle cases where d_type is DT_UNKNOWN
				if (pDirent->d_type == DT_UNKNOWN)
				{
					string sPath = sParent + "/" + sName;
					Stat st;
					if (UStat(sPath.c_str(), &st) == 0)
					{
						if (S_ISREG(st.st_mode))
						{
							pDirent->d_type = DT_REG;
						}
						else if (S_ISDIR(st.st_mode))
						{
							pDirent->d_type = DT_DIR;
						}
					}
				}
				if (pDirent->d_type == DT_DIR && strcmp(pDirent->d_name, ".") != 0 && strcmp(pDirent->d_name, "..") != 0)
				{
					string sDir = sParent + "/" + pDirent->d_name;
					if (nDepth == nDepthMax)
					{
						if (!mDir.insert(make_pair(UToU8(pDirent->d_name), sDir).second))
						{
							return 1;
						}
					}
					if (nDepth + 1 <= nDepthMax)
					{
						qDir.push(make_pair(sDir, nDepth + 1));
					}
				}
			}
			closedir(pDir);
		}
#endif
		qDir.pop();
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
	string sList = pTemp;
	delete[] pTemp;
	vector<string> vList = SplitOf(sList, "\r\n");
	for (vector<string>::iterator it = vList.begin(); it != vList.end(); ++it)
	{
		*it = trim(*it);
	}
	vector<string>::iterator itList = remove_if(vList.begin(), vList.end(), empty);
	vList.erase(itList, vList.end());
	vector<SRecord> vRecord;
	for (vector<string>::const_iterator it = vList.begin(); it != vList.end(); ++it)
	{
		const string& sLine = *it;
		SRecord record;
		record.Year = SToN32(sLine.substr(0, 4));
		record.Name = sLine.substr(5);
		map<string, UString>::iterator itDir = mDir.find(record.Name);
		if (itDir != mDir.end())
		{
			record.Path = UToU8(itDir->second);
			record.Exist = true;
			mDir.erase(itDir);
		}
		else
		{
			record.Exist = false;
		}
		vRecord.push_back(record);
	}
	for (map<string, UString>::const_iterator it = mDir.begin(); it != mDir.end(); ++it)
	{
		SRecord record;
		record.Year = 0;
		record.Name = it->first;
		record.Path = UToU8(it->second);
		record.Exist = true;
		vRecord.push_back(record);
	}
	fp = UFopen(argv[4], USTR("wb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	fprintf(fp, "[yes]\t[no]\t[year]\t[name]\t[path]\r\n");
	n32 nYear = -1;
	for (vector<SRecord>::const_iterator it = vRecord.begin(); it != vRecord.end(); ++it)
	{
		const SRecord& record = *it;
		if (record.Year != nYear)
		{
			fprintf(fp, "\r\n");
			nYear = record.Year;
		}
		if (record.Exist)
		{
			fprintf(fp, "[v]\t   \t");
		}
		else
		{
			fprintf(fp, "   \t[x]\t");
		}
		fprintf(fp, "%04d\t", record.Year);
		fprintf(fp, "%s\t", record.Name.c_str());
		fprintf(fp, "%s\r\n", record.Path.c_str());
	}
	fclose(fp);
	return 0;
}
