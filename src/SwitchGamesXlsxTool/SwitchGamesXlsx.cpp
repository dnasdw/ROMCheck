#include "SwitchGamesXlsx.h"
#include <tinyxml2.h>

CSwitchGamesXlsx::STextFileContent::STextFileContent()
	: EncodingOld(kEncodingUnknown)
	, EncodingNew(kEncodingUnknown)
	, LineTypeOld(kLineTypeUnknown)
	, LineTypeNew(kLineTypeUnknown)
{
}

CSwitchGamesXlsx::CSwitchGamesXlsx()
	: m_bStyleIsGreen(true)
	, m_bResave(false)
	, m_bCompact(false)
	, m_nActiveTabOld(0)
	, m_nActiveTabNew(0)
	, m_nSstCount(-1)
{
	m_sModuleDirName = UGetModuleDirName();
}

CSwitchGamesXlsx::~CSwitchGamesXlsx()
{
}

void CSwitchGamesXlsx::SetXlsxDirName(const UString& a_sXlsxDirName)
{
	m_sXlsxDirName = a_sXlsxDirName;
}

void CSwitchGamesXlsx::SetTableDirName(const UString& a_sTableDirName)
{
	m_sTableDirName = a_sTableDirName;
}

void CSwitchGamesXlsx::SetResultFileName(const UString& a_sResultFileName)
{
	m_sResultFileName = a_sResultFileName;
}

void CSwitchGamesXlsx::SetRemoteDirName(const UString& a_sRemoteDirName)
{
	m_sRemoteDirName = a_sRemoteDirName;
}

void CSwitchGamesXlsx::SetBaiduUserId(const UString& a_sBaiduUserId)
{
	m_sBaiduUserId = a_sBaiduUserId;
}

void CSwitchGamesXlsx::SetStyleIsGreen(bool a_bStyleIsGreen)
{
	m_bStyleIsGreen = a_bStyleIsGreen;
}

int CSwitchGamesXlsx::Resave()
{
	m_bResave = true;
	if (readConfig() != 0)
	{
		return 1;
	}
	// read /xl/workbook.xml
	if (readWorkbook() != 0)
	{
		return 1;
	}
	// read /xl/sharedStrings.xml
	if (readSharedStrings() != 0)
	{
		return 1;
	}
	// read /xl/styles.xml
	if (readStyles() != 0)
	{
		return 1;
	}
	// /_rels/.rels
	if (resaveRels() != 0)
	{
		return 1;
	}
	// /docProps/app.xml
	if (resaveApp() != 0)
	{
		return 1;
	}
	// /docProps/core.xml
	if (resaveCore() != 0)
	{
		return 1;
	}
	// /xl/_rels/workbook.xml.rels
	if (resaveWorkbookRels() != 0)
	{
		return 1;
	}
	// /xl/theme/theme1.xml
	if (resaveTheme1() != 0)
	{
		return 1;
	}
	// read /xl/worksheets/sheet%d.xml
	if (readSheet() != 0)
	{
		return 1;
	}
	updateSharedStrings();
	// write /xl/worksheets/sheet%d.xml
	if (writeSheet() != 0)
	{
		return 1;
	}
	// write /xl/sharedStrings.xml
	if (writeSharedStrings() != 0)
	{
		return 1;
	}
	// write /xl/styles.xml
	if (resaveStyles() != 0)
	{
		return 1;
	}
	// write /xl/workbook.xml
	if (resaveWorkbook() != 0)
	{
		return 1;
	}
	// /[Content_Types].xml
	if (resaveContentTypes() != 0)
	{
		return 1;
	}
	return 0;
}

int CSwitchGamesXlsx::Export()
{
	if (readConfig() != 0)
	{
		return 1;
	}
	// read /xl/workbook.xml
	if (readWorkbook() != 0)
	{
		return 1;
	}
	// read /xl/sharedStrings.xml
	if (readSharedStrings() != 0)
	{
		return 1;
	}
	// /xl/worksheets/sheet%d.xml
	if (readSheet() != 0)
	{
		return 1;
	}
	updateSharedStrings();
	if (writeTable() != 0)
	{
		return 1;
	}
	return 0;
}

int CSwitchGamesXlsx::Import()
{
	if (readConfig() != 0)
	{
		return 1;
	}
	if (readTable() != 0)
	{
		return 1;
	}
	updateSharedStrings();
	// write /xl/worksheets/sheet%d.xml
	if (writeSheet() != 0)
	{
		return 1;
	}
	// write /xl/sharedStrings.xml
	if (writeSharedStrings() != 0)
	{
		return 1;
	}
	return 0;
}

int CSwitchGamesXlsx::Sort()
{
	if (readConfig() != 0)
	{
		return 1;
	}
	if (readTable() != 0)
	{
		return 1;
	}
	if (sortTable() != 0)
	{
		return 1;
	}
	if (writeTable() != 0)
	{
		return 1;
	}
	return 0;
}

int CSwitchGamesXlsx::Check()
{
	if (readConfig() != 0)
	{
		return 1;
	}
	if (readTable() != 0)
	{
		return 1;
	}
	if (readResult() != 0)
	{
		return 1;
	}
	if (checkTable() != 0)
	{
		return 1;
	}
	updateSharedStrings();
	if (writeTable() != 0)
	{
		return 1;
	}
	return 0;
}

int CSwitchGamesXlsx::MakeRclonePatchBat()
{
	if (readConfig() != 0)
	{
		return 1;
	}
	if (readTable() != 0)
	{
		return 1;
	}
	if (makePatchTypeFileList() != 0)
	{
		return 1;
	}
	if (makeRclonePatchBat() != 0)
	{
		return 1;
	}
	return 0;
}

int CSwitchGamesXlsx::MakeBaiduPCSGoPatchBat()
{
	if (readConfig() != 0)
	{
		return 1;
	}
	if (readTable() != 0)
	{
		return 1;
	}
	if (makePatchTypeFileList() != 0)
	{
		return 1;
	}
	if (makeBaiduPCSGoPatchBat() != 0)
	{
		return 1;
	}
	return 0;
}

string CSwitchGamesXlsx::validateText(const string& a_sText)
{
	string sText = Replace(a_sText, '&', "&amp;");
	sText = Replace(sText, '<', "&lt;");
	sText = Replace(sText, '>', "&gt;");
	return sText;
}

string CSwitchGamesXlsx::validateValue(const string& a_sValue)
{
	string sValue = Replace(a_sValue, '\"', "&quot;");
	return sValue;
}

bool CSwitchGamesXlsx::travelElement(const tinyxml2::XMLElement* a_pRootElement, string& a_sXml)
{
	a_sXml += "<";
	a_sXml += a_pRootElement->Name();
	for (const tinyxml2::XMLAttribute* pRootElementAttribute = a_pRootElement->FirstAttribute(); pRootElementAttribute != nullptr; pRootElementAttribute = pRootElementAttribute->Next())
	{
		string sAttributeName = pRootElementAttribute->Name();
		string sAttributeValue = validateValue(pRootElementAttribute->Value());
		a_sXml += " " + sAttributeName + "=\"" + sAttributeValue + "\"";
	}
	const tinyxml2::XMLNode* pRootElementChild = a_pRootElement->FirstChild();
	if (pRootElementChild == nullptr)
	{
		a_sXml += "/>";
	}
	else
	{
		a_sXml += ">";
		for (; pRootElementChild != nullptr; pRootElementChild = pRootElementChild->NextSibling())
		{
			if (pRootElementChild->ToText() != nullptr)
			{
				a_sXml += validateText(pRootElementChild->Value());
			}
			else if (pRootElementChild->ToElement() != nullptr)
			{
				const tinyxml2::XMLElement* pRootElementChildElement = pRootElementChild->ToElement();
				if (!travelElement(pRootElementChildElement, a_sXml))
				{
					return false;
				}
			}
			else
			{
				return false;
			}
		}
		a_sXml += "</";
		a_sXml += a_pRootElement->Name();
		a_sXml += ">";
	}
	return true;
}

string CSwitchGamesXlsx::encodePosition(n32 a_nRowIndex, n32 a_nColumnIndex)
{
	string sPosition;
	n32 nColumnIndex = a_nColumnIndex + 1;
	while (nColumnIndex != 0)
	{
		n32 nQuot = nColumnIndex / 26;
		n32 nRem = nColumnIndex % 26;
		if (nRem == 0)
		{
			nRem += 26;
			nQuot--;
		}
		sPosition = static_cast<char>('A' + nRem - 1) + sPosition;
		nColumnIndex = nQuot;
	}
	sPosition += Format("%d", a_nRowIndex + 1);
	return sPosition;
}

bool CSwitchGamesXlsx::copyFile(const UString& a_sDestFileName, const UString& a_sSrcFileName)
{
	FILE* fpSrc = UFopen(a_sSrcFileName.c_str(), USTR("rb"), false);
	if (fpSrc == nullptr)
	{
		return false;
	}
	fseek(fpSrc, 0, SEEK_END);
	n64 nFileSize = ftell(fpSrc);
	FILE* fpDest = UFopen(a_sDestFileName.c_str(), USTR("wb"), false);
	if (fpDest == nullptr)
	{
		fclose(fpSrc);
		return false;
	}
	CopyFile(fpDest, fpSrc, 0, nFileSize);
	fclose(fpDest);
	fclose(fpSrc);
	return true;
}

string CSwitchGamesXlsx::trim(const string& a_sLine)
{
	return WToU8(trim(U8ToW(a_sLine)));
}

wstring CSwitchGamesXlsx::trim(const wstring& a_sLine)
{
	wstring sTrimmed = a_sLine;
	wstring::size_type uPos = sTrimmed.find_first_not_of(L"\n\v\f\r \x85\xA0");
	if (uPos == wstring::npos)
	{
		return L"";
	}
	sTrimmed.erase(0, uPos);
	uPos = sTrimmed.find_last_not_of(L"\n\v\f\r \x85\xA0");
	if (uPos != wstring::npos)
	{
		sTrimmed.erase(uPos + 1);
	}
	return sTrimmed;
}

bool CSwitchGamesXlsx::empty(const string& a_sLine)
{
	return trim(a_sLine).empty();
}

bool CSwitchGamesXlsx::pathCompare(const UString& lhs, const UString& rhs)
{
	wstring sLhs = UToW(lhs);
	transform(sLhs.begin(), sLhs.end(), sLhs.begin(), ::towupper);
	wstring sRhs = UToW(rhs);
	transform(sRhs.begin(), sRhs.end(), sRhs.begin(), ::towupper);
	return sLhs < sRhs;
}

bool CSwitchGamesXlsx::rowColumnTextCompare(const pair<n32, wstring>& lhs, const pair<n32, wstring>& rhs)
{
	return pathCompare(WToU(lhs.second), WToU(rhs.second));
}

bool CSwitchGamesXlsx::fileListCompare(const pair<UString, bool>& lhs, const pair<UString, bool>& rhs)
{
	return pathCompare(lhs.first, rhs.first);
}

int CSwitchGamesXlsx::readTextFile(const UString& a_sFilePath, STextFileContent& a_TextFileContent)
{
	FILE* fp = UFopen(a_sFilePath.c_str(), USTR("rb"), true);
	if (fp == nullptr)
	{
		return 1;
	}
	fseek(fp, 0, SEEK_END);
	u32 uFileSize = ftell(fp);
	if (uFileSize == 0)
	{
		fclose(fp);
		UPrintf(USTR("file size == 0: %") PRIUS USTR("\n"), a_sFilePath.c_str());
		return 1;
	}
	fseek(fp, 0, SEEK_SET);
	char* pTemp = new char[uFileSize + 1];
	fread(pTemp, 1, uFileSize, fp);
	fclose(fp);
	pTemp[uFileSize] = 0;
	a_TextFileContent.TextOld = pTemp;
	if (strlen(pTemp) != uFileSize || a_TextFileContent.TextOld.size() != uFileSize || memcmp(a_TextFileContent.TextOld.c_str(), pTemp, uFileSize) != 0)
	{
		delete[] pTemp;
		UPrintf(USTR("unicode?: %") PRIUS USTR("\n"), a_sFilePath.c_str());
		return 1;
	}
	delete[] pTemp;
	bool bCP437 = false;
	bool bUTF8 = false;
	if (StartWith(a_TextFileContent.TextOld, "\xEF\xBB\xBF"))
	{
		try
		{
			wstring sTextW = U8ToW(a_TextFileContent.TextOld);
			a_TextFileContent.TextNew = WToU8(sTextW);
			if (a_TextFileContent.TextNew == a_TextFileContent.TextOld)
			{
				bUTF8 = true;
				a_TextFileContent.EncodingOld = kEncodingUTF8withBOM;
			}
		}
		catch (...)
		{
			// do nothing
		}
	}
	if (!bUTF8)
	{
		try
		{
			wstring sTextW = XToW(a_TextFileContent.TextOld, 437, "CP437");
			a_TextFileContent.TextNew = WToX(sTextW, 437, "CP437");
			if (a_TextFileContent.TextNew == a_TextFileContent.TextOld)
			{
				bCP437 = true;
				a_TextFileContent.EncodingOld = kEncodingCP437;
			}
		}
		catch (...)
		{
			// do nothing
		}
	}
	if (!bCP437 && !bUTF8)
	{
		try
		{
			wstring sTextW = U8ToW(a_TextFileContent.TextOld);
			a_TextFileContent.TextNew = WToU8(sTextW);
			if (a_TextFileContent.TextNew == a_TextFileContent.TextOld)
			{
				bUTF8 = true;
				if (StartWith(a_TextFileContent.TextOld, "\xEF\xBB\xBF"))
				{
					a_TextFileContent.EncodingOld = kEncodingUTF8withBOM;
				}
				else
				{
					a_TextFileContent.EncodingOld = kEncodingUTF8;
				}
			}
		}
		catch (...)
		{
			// do nothing
		}
	}
	if (bCP437)
	{
		a_TextFileContent.TextNew = a_TextFileContent.TextOld;
		a_TextFileContent.EncodingNew = kEncodingCP437;
	}
	else if (bUTF8)
	{
		string sTextU8 = a_TextFileContent.TextOld;
		if (a_TextFileContent.EncodingOld == kEncodingUTF8withBOM)
		{
			sTextU8.erase(0, 3);
		}
		bool bUTF8ToCP437 = false;
		try
		{
			wstring sTextW = U8ToW(sTextU8);
			string sTextX = WToX(sTextW, 437, "CP437");
			wstring sTextXW = XToW(sTextX, 437, "CP437");
			if (sTextXW == sTextW)
			{
				a_TextFileContent.TextNew = sTextX;
				a_TextFileContent.EncodingNew = kEncodingCP437;
				bUTF8ToCP437 = true;
			}
		}
		catch (...)
		{
			// do nothing
		}
		if (!bUTF8ToCP437)
		{
			a_TextFileContent.TextNew = a_TextFileContent.TextOld;
			if (a_TextFileContent.EncodingOld == kEncodingUTF8)
			{
				a_TextFileContent.TextNew.insert(0, "\xEF\xBB\xBF");
			}
			a_TextFileContent.EncodingNew = kEncodingUTF8withBOM;
		}
	}
	else
	{
		a_TextFileContent.TextNew = a_TextFileContent.TextOld;
	}
	string sTextNoCRLF = Replace(a_TextFileContent.TextOld, "\r\n", "");
	string sTextNoCR = Replace(a_TextFileContent.TextOld, "\r", "");
	string sTextNoLF = Replace(a_TextFileContent.TextOld, "\n", "");
	n32 nCRLFCount = (a_TextFileContent.TextOld.size() - sTextNoCRLF.size()) / 2;
	n32 nCROnlyCount = (a_TextFileContent.TextOld.size() - sTextNoCR.size()) - nCRLFCount;
	n32 nLFOnlyCount = (a_TextFileContent.TextOld.size() - sTextNoLF.size()) - nCRLFCount;
	if (nCROnlyCount > nCRLFCount && nCROnlyCount > nLFOnlyCount)
	{
		if (nCRLFCount == 0 && nLFOnlyCount == 0)
		{
			a_TextFileContent.LineTypeOld = kLineTypeCR;
		}
		else
		{
			a_TextFileContent.LineTypeOld = kLineTypeCRMix;
		}
	}
	else if (nCRLFCount > nLFOnlyCount)
	{
		if (nLFOnlyCount == 0)
		{
			if (nCROnlyCount == 0)
			{
				a_TextFileContent.LineTypeOld = kLineTypeCRLF;
			}
			else
			{
				a_TextFileContent.LineTypeOld = kLineTypeCRLF_CR;
			}
		}
		else
		{
			a_TextFileContent.LineTypeOld = kLineTypeCRLFMix;
		}
	}
	else if (nLFOnlyCount > nCRLFCount)
	{
		if (nCRLFCount == 0)
		{
			if (nCROnlyCount == 0)
			{
				a_TextFileContent.LineTypeOld = kLineTypeLF;
			}
			else
			{
				a_TextFileContent.LineTypeOld = kLineTypeLF_CR;
			}
		}
		else
		{
			a_TextFileContent.LineTypeOld = kLineTypeLFMix;
		}
	}
	if (EndWith(a_sFilePath, USTR(".nfo")) && a_TextFileContent.LineTypeOld != kLineTypeUnknown && nCROnlyCount == 0)
	{
		a_TextFileContent.TextNew = Replace(a_TextFileContent.TextNew, "\r", "");
		a_TextFileContent.LineTypeNew = kLineTypeLF;
	}
	else if (EndWith(a_sFilePath, USTR(".nfo")) && a_TextFileContent.LineTypeOld == kLineTypeCR)
	{
		a_TextFileContent.TextNew = Replace(a_TextFileContent.TextNew, "\r", "\n");
		a_TextFileContent.LineTypeNew = kLineTypeLF;
	}
	else if (EndWith(a_sFilePath, USTR(".sfv")) && a_TextFileContent.LineTypeOld != kLineTypeUnknown && nCROnlyCount == 0)
	{
		a_TextFileContent.TextNew = Replace(a_TextFileContent.TextNew, "\r", "");
		a_TextFileContent.TextNew = Replace(a_TextFileContent.TextNew, "\n", "\r\n");
		a_TextFileContent.LineTypeNew = kLineTypeCRLF;
	}
	else if (EndWith(a_sFilePath, USTR(".sfv")) && a_TextFileContent.LineTypeOld == kLineTypeCR)
	{
		a_TextFileContent.TextNew = Replace(a_TextFileContent.TextNew, "\r", "\r\n");
		a_TextFileContent.LineTypeNew = kLineTypeCRLF;
	}
	else
	{
		a_TextFileContent.LineTypeNew = a_TextFileContent.LineTypeOld;
	}
	return 0;
}

bool CSwitchGamesXlsx::makeDir(const UString& a_sDirPath)
{
	UString sDirPath = a_sDirPath;
#if SDW_PLATFORM == SDW_PLATFORM_WINDOWS
	u32 uMaxPath = sDirPath.size() + MAX_PATH * 2;
	wchar_t* pDirPath = new wchar_t[uMaxPath];
	if (_wfullpath(pDirPath, UToW(sDirPath).c_str(), uMaxPath) == nullptr)
	{
		return false;
	}
	sDirPath = WToU(pDirPath);
	delete[] pDirPath;
	if (!StartWith(sDirPath, USTR("\\\\")))
	{
		sDirPath = USTR("\\\\?\\") + sDirPath;
	}
#endif
	UString sPrefix;
	if (StartWith(sDirPath, USTR("\\\\?\\")))
	{
		sPrefix = USTR("\\\\?\\");
		sDirPath.erase(0, 4);
	}
	else if (StartWith(sDirPath, USTR("\\\\")))
	{
		sPrefix = USTR("\\\\");
		sDirPath.erase(0, 2);
	}
	vector<UString> vDirPath = SplitOf(sDirPath, USTR("/\\"));
	UString sDirName = sPrefix;
	UString sSep = sPrefix.empty() ? USTR("/") : USTR("\\");
	for (n32 i = 0; i < static_cast<n32>(vDirPath.size()); i++)
	{
		sDirName += vDirPath[i];
		if (!sPrefix.empty() && i < 1)
		{
			// do nothing
		}
		else if (!UMakeDir(sDirName.c_str()))
		{
			return false;
		}
		sDirName += sSep;
	}
	return true;
}

int CSwitchGamesXlsx::writeFileString(const UString& a_sFilePath, const string& a_sStringContent)
{
	UString sFilePath = a_sFilePath;
	UString sDirPath = USTR(".");
	UString::size_type uPos = a_sFilePath.find_last_of(USTR("/\\"));
	if (uPos != UString::npos)
	{
		sDirPath = a_sFilePath.substr(0, uPos);
	}
	if (!makeDir(sDirPath))
	{
		return 1;
	}
#if SDW_PLATFORM == SDW_PLATFORM_WINDOWS
	u32 uMaxPath = sDirPath.size() + MAX_PATH * 2;
	wchar_t* pFilePath = new wchar_t[uMaxPath];
	if (_wfullpath(pFilePath, UToW(a_sFilePath).c_str(), uMaxPath) == nullptr)
	{
		return 1;
	}
	sFilePath = WToU(pFilePath);
	delete[] pFilePath;
	if (!StartWith(sFilePath, USTR("\\\\")))
	{
		sFilePath = USTR("\\\\?\\") + sFilePath;
	}
#endif
	FILE* fp = UFopen(sFilePath.c_str(), USTR("wb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	fwrite(a_sStringContent.c_str(), 1, a_sStringContent.size(), fp);
	fclose(fp);
	return 0;
}

int CSwitchGamesXlsx::readConfig()
{
	UString sConfigXmlPath = m_sModuleDirName + USTR("/Switch Games.xml");
	FILE* fp = UFopen(sConfigXmlPath.c_str(), USTR("rb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	tinyxml2::XMLDocument xmlDoc;
	tinyxml2::XMLError eError = xmlDoc.LoadFile(fp);
	fclose(fp);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		return 1;
	}
	const tinyxml2::XMLElement* pDocConfig = xmlDoc.FirstChildElement("config");
	if (pDocConfig == nullptr)
	{
		return 1;
	}
	const tinyxml2::XMLElement* pConfigSheets = pDocConfig->FirstChildElement("sheets");
	if (pConfigSheets == nullptr)
	{
		return 1;
	}
	for (const tinyxml2::XMLElement* pSheetsSheet = pConfigSheets->FirstChildElement("sheet"); pSheetsSheet != nullptr; pSheetsSheet = pSheetsSheet->NextSiblingElement("sheet"))
	{
		wstring sSheetText;
		const char* pSheetText = pSheetsSheet->GetText();
		if (pSheetText != nullptr)
		{
			sSheetText = U8ToW(pSheetText);
		}
		m_vSheetName.push_back(sSheetText);
		if (m_bResave)
		{
			m_sSheetName.insert(sSheetText);
		}
	}
	const tinyxml2::XMLElement* pConfigGreen = pDocConfig->FirstChildElement("green");
	if (pConfigGreen == nullptr)
	{
		return 1;
	}
	for (const tinyxml2::XMLElement* pGreenName = pConfigGreen->FirstChildElement("name"); pGreenName != nullptr; pGreenName = pGreenName->NextSiblingElement("name"))
	{
		wstring sNameText;
		const char* pNameText = pGreenName->GetText();
		if (pNameText != nullptr)
		{
			sNameText = U8ToW(pNameText);
		}
		if (sNameText.empty())
		{
			continue;
		}
		try
		{
			wregex rPattern(sNameText, regex_constants::ECMAScript | regex_constants::icase);
			m_vGreenStyleNamePattern.push_back(rPattern);
		}
		catch (regex_error& e)
		{
			UPrintf(USTR("%") PRIUS USTR(" regex error: %") PRIUS USTR("\n\n"), WToU(sNameText).c_str(), AToU(e.what()).c_str());
			continue;
		}
	}
	const tinyxml2::XMLElement* pGreenExclude = pConfigGreen->FirstChildElement("exclude");
	if (pGreenExclude != nullptr)
	{
		for (const tinyxml2::XMLElement* pExcludeName = pGreenExclude->FirstChildElement("name"); pExcludeName != nullptr; pExcludeName = pExcludeName->NextSiblingElement("name"))
		{
			wstring sNameText;
			const char* pNameText = pExcludeName->GetText();
			if (pNameText != nullptr)
			{
				sNameText = U8ToW(pNameText);
			}
			if (sNameText.empty())
			{
				continue;
			}
			try
			{
				wregex rPattern(sNameText, regex_constants::ECMAScript | regex_constants::icase);
				m_vNotGreenStyleNamePattern.push_back(rPattern);
			}
			catch (regex_error& e)
			{
				UPrintf(USTR("%") PRIUS USTR(" regex error: %") PRIUS USTR("\n\n"), WToU(sNameText).c_str(), AToU(e.what()).c_str());
				continue;
			}
		}
	}
	return 0;
}

int CSwitchGamesXlsx::readWorkbook()
{
	UString sWorkbookXmlPath = m_sXlsxDirName + USTR("/xl/workbook.xml");
	FILE* fp = UFopen(sWorkbookXmlPath.c_str(), USTR("rb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	fseek(fp, 0, SEEK_END);
	u32 uWorkbookXmlSize = ftell(fp);
	fseek(fp, 0, SEEK_SET);
	char* pWorkbookXml = new char[uWorkbookXmlSize + 1];
	fread(pWorkbookXml, 1, uWorkbookXmlSize, fp);
	pWorkbookXml[uWorkbookXmlSize] = 0;
	string sWorkbookXml = pWorkbookXml;
	delete[] pWorkbookXml;
	tinyxml2::XMLDocument xmlDoc;
	tinyxml2::XMLError eError = xmlDoc.LoadFile(fp);
	fclose(fp);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		return 1;
	}
	if (!StartWith(sWorkbookXml, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"))
	{
		return 1;
	}
	const tinyxml2::XMLElement* pDocWorkbook = xmlDoc.RootElement();
	if (pDocWorkbook == nullptr)
	{
		return 1;
	}
	string sWorkbookName = pDocWorkbook->Name();
	if (sWorkbookName != "workbook")
	{
		return 1;
	}
	const tinyxml2::XMLElement* pWorkbookBookViews = pDocWorkbook->FirstChildElement("bookViews");
	if (pWorkbookBookViews != nullptr)
	{
		const tinyxml2::XMLElement* pBookViewsWorkbookView = pWorkbookBookViews->FirstChildElement("workbookView");
		if (pBookViewsWorkbookView != nullptr)
		{
			const char* pWorkbookViewActiveTab = pBookViewsWorkbookView->Attribute("activeTab");
			if (pWorkbookViewActiveTab != nullptr)
			{
				m_nActiveTabOld = SToN32(pWorkbookViewActiveTab);
			}
		}
	}
	if (!m_bResave)
	{
		m_nActiveTabNew = m_nActiveTabOld;
	}
	return 0;
}

int CSwitchGamesXlsx::readSharedStrings()
{
	UString sSharedStringsXmlPath = m_sXlsxDirName + USTR("/xl/sharedStrings.xml");
	FILE* fp = UFopen(sSharedStringsXmlPath.c_str(), USTR("rb"), false);
	if (fp == nullptr)
	{
		return 0;
	}
	fseek(fp, 0, SEEK_END);
	u32 uSharedStringsXmlSize = ftell(fp);
	fseek(fp, 0, SEEK_SET);
	char* pSharedStringsXml0 = new char[uSharedStringsXmlSize + 1];
	fread(pSharedStringsXml0, 1, uSharedStringsXmlSize, fp);
	pSharedStringsXml0[uSharedStringsXmlSize] = 0;
	string sSharedStringsXml0 = pSharedStringsXml0;
	delete[] pSharedStringsXml0;
	string sSharedStringsXml;
	sSharedStringsXml.reserve(uSharedStringsXmlSize * 2);
	fseek(fp, 0, SEEK_SET);
	tinyxml2::XMLDocument xmlDoc;
	tinyxml2::XMLError eError = xmlDoc.LoadFile(fp);
	fclose(fp);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		return 1;
	}
	if (!StartWith(sSharedStringsXml0, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"))
	{
		return 1;
	}
	sSharedStringsXml += "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
	const tinyxml2::XMLElement* pDocSst = xmlDoc.RootElement();
	if (pDocSst == nullptr)
	{
		return 1;
	}
	string sSstName = pDocSst->Name();
	if (sSstName != "sst")
	{
		return 1;
	}
	n32 nSstUniqueCount = -1;
	sSharedStringsXml += "<";
	sSharedStringsXml += pDocSst->Name();
	for (const tinyxml2::XMLAttribute* pSstAttribute = pDocSst->FirstAttribute(); pSstAttribute != nullptr; pSstAttribute = pSstAttribute->Next())
	{
		string sAttributeName = pSstAttribute->Name();
		string sAttributeValue = pSstAttribute->Value();
		if (sAttributeName == "count")
		{
			m_nSstCount = SToN32(sAttributeValue);
		}
		else if (sAttributeName == "uniqueCount")
		{
			nSstUniqueCount = SToN32(sAttributeValue);
		}
		sAttributeValue = validateValue(sAttributeValue);
		sSharedStringsXml += " " + sAttributeName + "=\"" + sAttributeValue + "\"";
	}
	sSharedStringsXml += ">";
	if (m_nSstCount < 0 || nSstUniqueCount < 0)
	{
		return 1;
	}
	n32 nSstSiCount = 0;
	for (const tinyxml2::XMLNode* pSstChild = pDocSst->FirstChild(); pSstChild != nullptr; pSstChild = pSstChild->NextSibling())
	{
		if (pSstChild->ToText() != nullptr)
		{
			return 1;
		}
		const tinyxml2::XMLElement* pSstChildElement = pSstChild->ToElement();
		if (pSstChildElement == nullptr)
		{
			return 1;
		}
		string sSstChildElementName = pSstChildElement->Name();
		if (sSstChildElementName != "si")
		{
			// no other child element
			return 1;
			if (!travelElement(pSstChildElement, sSharedStringsXml))
			{
				return 1;
			}
		}
		else
		{
			const tinyxml2::XMLElement* pSstSi = pSstChildElement;
			string sText;
			sSharedStringsXml += "<";
			sSharedStringsXml += pSstSi->Name();
			for (const tinyxml2::XMLAttribute* pSiAttribute = pSstSi->FirstAttribute(); pSiAttribute != nullptr; pSiAttribute = pSiAttribute->Next())
			{
				// no attribute
				return 1;
				string sAttributeName = pSiAttribute->Name();
				string sAttributeValue = validateValue(pSiAttribute->Value());
				sSharedStringsXml += " " + sAttributeName + "=\"" + sAttributeValue + "\"";
			}
			sSharedStringsXml += ">";
			n32 nSiTCount = 0;
			for (const tinyxml2::XMLNode* pSiChild = pSstSi->FirstChild(); pSiChild != nullptr; pSiChild = pSiChild->NextSibling())
			{
				if (pSiChild->ToText() != nullptr)
				{
					return 1;
				}
				const tinyxml2::XMLElement* pSiChildElement = pSiChild->ToElement();
				if (pSiChildElement == nullptr)
				{
					return 1;
				}
				string sSiChildElementName = pSiChildElement->Name();
				if (sSiChildElementName != "phoneticPr" && sSiChildElementName != "t")
				{
					// no other child element
					return 1;
					if (!travelElement(pSiChildElement, sSharedStringsXml))
					{
						return 1;
					}
				}
				else if (sSiChildElementName == "phoneticPr")
				{
					if (!travelElement(pSiChildElement, sSharedStringsXml))
					{
						return 1;
					}
					const tinyxml2::XMLElement* pSiPhoneticPr = pSiChildElement;
					for (const tinyxml2::XMLAttribute* pPhoneticPrAttribute = pSiPhoneticPr->FirstAttribute(); pPhoneticPrAttribute != nullptr; pPhoneticPrAttribute = pPhoneticPrAttribute->Next())
					{
						string sAttributeName = pPhoneticPrAttribute->Name();
						string sAttributeValue = pPhoneticPrAttribute->Value();
						if (sAttributeName == "fontId")
						{
							if (sAttributeValue != "1")
							{
								return 1;
							}
						}
						else if (sAttributeName == "type")
						{
							if (sAttributeValue != "noConversion")
							{
								return 1;
							}
						}
						else
						{
							return 1;
						}
					}
				}
				else if (sSiChildElementName == "t")
				{
					if (nSiTCount != 0)
					{
						return 1;
					}
					const tinyxml2::XMLElement* pSiT = pSiChildElement;
					bool bTPreserve = false;
					string sStmt;
					sSharedStringsXml += "<";
					sSharedStringsXml += pSiT->Name();
					for (const tinyxml2::XMLAttribute* pTAttribute = pSiT->FirstAttribute(); pTAttribute != nullptr; pTAttribute = pTAttribute->Next())
					{
						string sAttributeName = pTAttribute->Name();
						string sAttributeValue = pTAttribute->Value();
						if (sAttributeName != "xml:space" || sAttributeValue != "preserve")
						{
							return 1;
						}
						bTPreserve = true;
						sAttributeValue = validateValue(sAttributeValue);
						sSharedStringsXml += " " + sAttributeName + "=\"" + sAttributeValue + "\"";
					}
					sSharedStringsXml += ">";
					n32 nTChildCount = 0;
					for (const tinyxml2::XMLNode* pTChild = pSiT->FirstChild(); pTChild != nullptr; pTChild = pTChild->NextSibling())
					{
						const tinyxml2::XMLText* pTChildText = pTChild->ToText();
						if (pTChildText == nullptr)
						{
							return 1;
						}
						sStmt = pTChildText->Value();
						if (!bTPreserve)
						{
							sStmt = Trim(sStmt);
						}
						sSharedStringsXml += validateText(sStmt);
						sText += validateText(sStmt);
						nTChildCount++;
					}
					sSharedStringsXml += "</";
					sSharedStringsXml += pSiT->Name();
					sSharedStringsXml += ">";
					if (nTChildCount != 1)
					{
						return 1;
					}
					nSiTCount++;
				}
			}
			sSharedStringsXml += "</";
			sSharedStringsXml += pSstSi->Name();
			sSharedStringsXml += ">";
			if (!m_mSharedStrings.insert(make_pair(nSstSiCount, U8ToW(sText))).second)
			{
				return 1;
			}
			nSstSiCount++;
		}
	}
	sSharedStringsXml += "</";
	sSharedStringsXml += pDocSst->Name();
	sSharedStringsXml += ">";
	if (nSstSiCount != nSstUniqueCount)
	{
		return 1;
	}
	sSharedStringsXml = Replace(sSharedStringsXml, "\r\n", "\n");
	sSharedStringsXml = Replace(sSharedStringsXml, "\r", "");
	sSharedStringsXml = Replace(sSharedStringsXml, "\n", "\r\n");
	if (sSharedStringsXml != sSharedStringsXml0)
	{
		return 1;
	}
	if (m_bResave)
	{
		for (map<n32, wstring>::const_iterator it = m_mSharedStrings.begin(); it != m_mSharedStrings.end(); ++it)
		{
			const wstring& sText = it->second;
			m_mSharedStringsIndexNew.insert(make_pair(sText, -1));
		}
		m_mSharedStringsIndexNew[L"Release Name"] = 0;
		m_mSharedStringsIndexNew[L"File Name"] = 0;
		m_mSharedStringsIndexNew[L"Comment"] = 0;
		m_mSharedStringsIndexNew[L"Comment2"] = 0;
		m_mSharedStringsIndexNew[L".."] = 0;
		m_mSharedStringsIndexNew[L"nfo LF/sfv CRLF"] = 0;
	}
	return 0;
}

int CSwitchGamesXlsx::readStyles()
{
	UString sStylesXmlPath = m_sXlsxDirName + USTR("/xl/styles.xml");
	FILE* fp = UFopen(sStylesXmlPath.c_str(), USTR("rb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	fseek(fp, 0, SEEK_END);
	u32 uStylesXmlSize = ftell(fp);
	fseek(fp, 0, SEEK_SET);
	char* pStylesXml = new char[uStylesXmlSize + 1];
	fread(pStylesXml, 1, uStylesXmlSize, fp);
	pStylesXml[uStylesXmlSize] = 0;
	string sStylesXml = pStylesXml;
	delete[] pStylesXml;
	tinyxml2::XMLDocument xmlDoc;
	tinyxml2::XMLError eError = xmlDoc.LoadFile(fp);
	fclose(fp);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		return 1;
	}
	if (!StartWith(sStylesXml, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"))
	{
		return 1;
	}
	const tinyxml2::XMLElement* pDocStyleSheet = xmlDoc.RootElement();
	if (pDocStyleSheet == nullptr)
	{
		return 1;
	}
	string sStyleSheetName = pDocStyleSheet->Name();
	if (sStyleSheetName != "styleSheet")
	{
		return 1;
	}
	const tinyxml2::XMLElement* pStyleSheetFills = pDocStyleSheet->FirstChildElement("fills");
	if (pStyleSheetFills == nullptr)
	{
		return 1;
	}
	map<n32, n32> mFillOldToNew;
	n32 nFillIndexOld = 0;
	for (const tinyxml2::XMLElement* pFillsFill = pStyleSheetFills->FirstChildElement("fill"); pFillsFill != nullptr; pFillsFill = pFillsFill->NextSiblingElement("fill"))
	{
		const tinyxml2::XMLElement* pFillPatternFill = pFillsFill->FirstChildElement("patternFill");
		if (pFillPatternFill == nullptr)
		{
			return 1;
		}
		const char* pPatternFillPatternType = pFillPatternFill->Attribute("patternType");
		if (pPatternFillPatternType == nullptr)
		{
			return 1;
		}
		string sPatternFillPatternType = pPatternFillPatternType;
		if (sPatternFillPatternType == "none")
		{
			mFillOldToNew[nFillIndexOld++] = kFillIdNone;
		}
		else if (sPatternFillPatternType == "gray125")
		{
			mFillOldToNew[nFillIndexOld++] = kFillIdGray125;
		}
		else if (sPatternFillPatternType != "solid")
		{
			mFillOldToNew[nFillIndexOld++] = kFillIdRed;
			continue;
		}
		else
		{
			const tinyxml2::XMLElement* pPatternFillFgColor = pFillPatternFill->FirstChildElement("fgColor");
			if (pPatternFillFgColor == nullptr)
			{
				mFillOldToNew[nFillIndexOld++] = kFillIdRed;
				continue;
			}
			const char* pFgColorTheme = pPatternFillFgColor->Attribute("theme");
			const char* pFgColorRgb = pPatternFillFgColor->Attribute("rgb");
			if (pFgColorTheme == nullptr && pFgColorRgb == nullptr)
			{
				mFillOldToNew[nFillIndexOld++] = kFillIdRed;
				continue;
			}
			else if (pFgColorTheme != nullptr && pFgColorRgb != nullptr)
			{
				mFillOldToNew[nFillIndexOld++] = kFillIdRed;
				continue;
			}
			else if (pFgColorTheme != nullptr)
			{
				n32 nFgColorTheme = SToN32(pFgColorTheme);
				if (nFgColorTheme == 7)
				{
					mFillOldToNew[nFillIndexOld++] = kFillIdGold;
				}
				else if (nFgColorTheme == 9)
				{
					mFillOldToNew[nFillIndexOld++] = kFillIdGreen;
				}
				else
				{
					mFillOldToNew[nFillIndexOld++] = kFillIdRed;
				}
			}
			else if (pFgColorRgb != nullptr)
			{
				u32 uFgColorRgb = SToU32(pFgColorRgb, 16);
				if (uFgColorRgb == 0xFFFFC000)
				{
					mFillOldToNew[nFillIndexOld++] = kFillIdGold;
				}
				else if (uFgColorRgb == 0xFF70AD47)
				{
					mFillOldToNew[nFillIndexOld++] = kFillIdGreen;
				}
				else if (uFgColorRgb == 0xFFFFFF00)
				{
					mFillOldToNew[nFillIndexOld++] = kFillIdYellow;
				}
				else if (uFgColorRgb == 0xFFFF0000)
				{
					mFillOldToNew[nFillIndexOld++] = kFillIdRed;
				}
				else
				{
					mFillOldToNew[nFillIndexOld++] = kFillIdRed;
				}
			}
		}
	}
	const tinyxml2::XMLElement* pStyleSheetCellXfs = pDocStyleSheet->FirstChildElement("cellXfs");
	if (pStyleSheetCellXfs == nullptr)
	{
		return 1;
	}
	n32 nXfIndexOld = 0;
	for (const tinyxml2::XMLElement* pCellXfsXf = pStyleSheetCellXfs->FirstChildElement("xf"); pCellXfsXf != nullptr; pCellXfsXf = pCellXfsXf->NextSiblingElement("xf"))
	{
		n32 nXfFillId = kFillIdNone;
		bool bXfApplyFill = false;
		const char* pXfApplyFill = pCellXfsXf->Attribute("applyFill");
		if (pXfApplyFill != nullptr)
		{
			bXfApplyFill = SToN32(pXfApplyFill) == 1;
		}
		if (bXfApplyFill)
		{
			const char* pXfFillId = pCellXfsXf->Attribute("fillId");
			if (pXfFillId == nullptr)
			{
				nXfFillId = kFillIdRed;
			}
			else
			{
				nXfFillId = SToN32(pXfFillId);
				map<n32, n32>::iterator it = mFillOldToNew.find(nXfFillId);
				if (it == mFillOldToNew.end())
				{
					nXfFillId = kFillIdRed;
				}
				else
				{
					nXfFillId = it->second;
				}
			}
		}
		switch (nXfFillId)
		{
		case kFillIdNone:
			m_mStyleOldToNew[nXfIndexOld++] = kStyleIdNone;
			break;
		case kFillIdRed:
			m_mStyleOldToNew[nXfIndexOld++] = kStyleIdRed;
			break;
		case kFillIdBlue:
			m_mStyleOldToNew[nXfIndexOld++] = kStyleIdBlue;
			break;
		case kFillIdGreen:
			m_mStyleOldToNew[nXfIndexOld++] = kStyleIdGreen;
			break;
		case kFillIdYellow:
			m_mStyleOldToNew[nXfIndexOld++] = kStyleIdYellow;
			break;
		case kFillIdGold:
			m_mStyleOldToNew[nXfIndexOld++] = kStyleIdGold;
			break;
		default:
			m_mStyleOldToNew[nXfIndexOld++] = kStyleIdRed;
			break;
		}
	}
	return 0;
}

int CSwitchGamesXlsx::resaveRels() const
{
	if (!m_bResave)
	{
		return 1;
	}
	UString sRelsPath = m_sXlsxDirName + USTR("/_rels/.rels");
	UString sSwitchGamesRelsPath = m_sModuleDirName + USTR("/Switch Games/_rels/.rels");
	if (!copyFile(sRelsPath, sSwitchGamesRelsPath))
	{
		return 1;
	}
	return 0;
}

int CSwitchGamesXlsx::resaveApp()
{
	if (!m_bResave)
	{
		return 1;
	}
	UString sAppXmlPath = m_sXlsxDirName + USTR("/docProps/app.xml");
	FILE* fp = UFopen(sAppXmlPath.c_str(), USTR("rb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	fseek(fp, 0, SEEK_END);
	u32 uAppXmlSize = ftell(fp);
	fseek(fp, 0, SEEK_SET);
	char* pAppXml = new char[uAppXmlSize + 1];
	fread(pAppXml, 1, uAppXmlSize, fp);
	pAppXml[uAppXmlSize] = 0;
	string sAppXml = pAppXml;
	delete[] pAppXml;
	tinyxml2::XMLDocument xmlDoc;
	tinyxml2::XMLError eError = xmlDoc.LoadFile(fp);
	fclose(fp);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		return 1;
	}
	if (!StartWith(sAppXml, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"))
	{
		return 1;
	}
	tinyxml2::XMLElement* pDocProperties = xmlDoc.RootElement();
	if (pDocProperties == nullptr)
	{
		return 1;
	}
	string sPropertiesName = pDocProperties->Name();
	if (sPropertiesName != "Properties")
	{
		return 1;
	}
	tinyxml2::XMLElement* pPropertiesHeadingPairs = pDocProperties->FirstChildElement("HeadingPairs");
	if (pPropertiesHeadingPairs == nullptr)
	{
		return 1;
	}
	tinyxml2::XMLElement* pHeadingPairsVtVector = pPropertiesHeadingPairs->FirstChildElement("vt:vector");
	if (pHeadingPairsVtVector == nullptr)
	{
		return 1;
	}
	tinyxml2::XMLElement* pVtVectorVtVariant = pHeadingPairsVtVector->FirstChildElement("vt:variant");
	if (pVtVectorVtVariant == nullptr)
	{
		return 1;
	}
	pVtVectorVtVariant = pVtVectorVtVariant->NextSiblingElement("vt:variant");
	if (pVtVectorVtVariant == nullptr)
	{
		return 1;
	}
	tinyxml2::XMLElement* pVtVariantVtI4 = pVtVectorVtVariant->FirstChildElement("vt:i4");
	if (pVtVariantVtI4 == nullptr)
	{
		return 1;
	}
	n32 nSheetCountOld = pVtVariantVtI4->IntText();
	tinyxml2::XMLElement* pPropertiesTitlesOfParts = pDocProperties->FirstChildElement("TitlesOfParts");
	if (pPropertiesTitlesOfParts == nullptr)
	{
		return 1;
	}
	tinyxml2::XMLElement* pTitlesOfPartsVtVector = pPropertiesTitlesOfParts->FirstChildElement("vt:vector");
	if (pTitlesOfPartsVtVector == nullptr)
	{
		return 1;
	}
	if (pTitlesOfPartsVtVector->IntAttribute("size") != nSheetCountOld)
	{
		return 1;
	}
	n32 nSheetIndexOld = 1;
	for (tinyxml2::XMLElement* pVtVectorVtLpstr = pTitlesOfPartsVtVector->FirstChildElement("vt:lpstr"); pVtVectorVtLpstr != nullptr; pVtVectorVtLpstr = pVtVectorVtLpstr->NextSiblingElement("vt:lpstr"))
	{
		wstring sVtLpstrText;
		const char* pVtLpstrText = pVtVectorVtLpstr->GetText();
		if (pVtLpstrText != nullptr)
		{
			sVtLpstrText = U8ToW(pVtLpstrText);
		}
		if (m_sSheetName.find(sVtLpstrText) == m_sSheetName.end())
		{
			continue;
		}
		if (nSheetIndexOld == m_nActiveTabOld + 1)
		{
			for (n32 i = 0; i < static_cast<n32>(m_vSheetName.size()); i++)
			{
				if (m_vSheetName[i] == sVtLpstrText)
				{
					m_nActiveTabNew = i;
				}
			}
		}
		m_mSheetIndexOld[sVtLpstrText] = nSheetIndexOld++;
	}
	UString sSwitchGamesAppXmlPath = m_sModuleDirName + USTR("/Switch Games/docProps/app.xml");
	fp = UFopen(sSwitchGamesAppXmlPath.c_str(), USTR("rb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	eError = xmlDoc.LoadFile(fp);
	fclose(fp);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		return 1;
	}
	pDocProperties = xmlDoc.RootElement();
	pPropertiesHeadingPairs = pDocProperties->FirstChildElement("HeadingPairs");
	pHeadingPairsVtVector = pPropertiesHeadingPairs->FirstChildElement("vt:vector");
	pVtVectorVtVariant = pHeadingPairsVtVector->FirstChildElement("vt:variant");
	pVtVectorVtVariant = pVtVectorVtVariant->NextSiblingElement("vt:variant");
	pVtVariantVtI4 = pVtVectorVtVariant->FirstChildElement("vt:i4");
	pVtVariantVtI4->SetText(static_cast<n32>(m_vSheetName.size()));
	pPropertiesTitlesOfParts = pDocProperties->FirstChildElement("TitlesOfParts");
	pTitlesOfPartsVtVector = pPropertiesTitlesOfParts->FirstChildElement("vt:vector");
	pTitlesOfPartsVtVector->SetAttribute("size", static_cast<n32>(m_vSheetName.size()));
	pTitlesOfPartsVtVector->DeleteChildren();
	for (n32 i = 0; i < static_cast<n32>(m_vSheetName.size()); i++)
	{
		tinyxml2::XMLElement* pVtVectorVtLpstr = pTitlesOfPartsVtVector->InsertNewChildElement("vt:lpstr");
		pVtVectorVtLpstr->SetText(WToU8(m_vSheetName[i]).c_str());
	}
	fp = UFopen(sAppXmlPath.c_str(), USTR("wb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	eError = xmlDoc.SaveFile(fp, m_bCompact);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		fclose(fp);
		return 1;
	}
	fclose(fp);
	return 0;
}

int CSwitchGamesXlsx::resaveCore()
{
	if (!m_bResave)
	{
		return 1;
	}
	UString sCoreXmlPath = m_sXlsxDirName + USTR("/docProps/core.xml");
	FILE* fp = UFopen(sCoreXmlPath.c_str(), USTR("rb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	fseek(fp, 0, SEEK_END);
	u32 uCoreXmlSize = ftell(fp);
	fseek(fp, 0, SEEK_SET);
	char* pCoreXml = new char[uCoreXmlSize + 1];
	fread(pCoreXml, 1, uCoreXmlSize, fp);
	pCoreXml[uCoreXmlSize] = 0;
	string sCoreXml = pCoreXml;
	delete[] pCoreXml;
	tinyxml2::XMLDocument xmlDoc;
	tinyxml2::XMLError eError = xmlDoc.LoadFile(fp);
	fclose(fp);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		return 1;
	}
	if (!StartWith(sCoreXml, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"))
	{
		return 1;
	}
	tinyxml2::XMLElement* pDocCpCoreProperties = xmlDoc.RootElement();
	if (pDocCpCoreProperties == nullptr)
	{
		return 1;
	}
	string sCpCorePropertiesName = pDocCpCoreProperties->Name();
	if (sCpCorePropertiesName != "cp:coreProperties")
	{
		return 1;
	}
	tinyxml2::XMLElement* pCpCorePropertiesDcCreator = pDocCpCoreProperties->FirstChildElement("dc:creator");
	if (pCpCorePropertiesDcCreator == nullptr)
	{
		return 1;
	}
	const char* pDcCreatorText = pCpCorePropertiesDcCreator->GetText();
	if (pDcCreatorText != nullptr)
	{
		m_sCreator = pDcCreatorText;
	}
	tinyxml2::XMLElement* pCpCorePropertiesCpLastModifiedBy = pDocCpCoreProperties->FirstChildElement("cp:lastModifiedBy");
	if (pCpCorePropertiesCpLastModifiedBy == nullptr)
	{
		return 1;
	}
	const char* pCpCorePropertiesCpLastModifiedByText = pCpCorePropertiesCpLastModifiedBy->GetText();
	if (pCpCorePropertiesCpLastModifiedByText != nullptr)
	{
		m_sLastModifiedBy = pCpCorePropertiesCpLastModifiedByText;
	}
	tinyxml2::XMLElement* pCpCorePropertiesDctermsModified = pDocCpCoreProperties->FirstChildElement("dcterms:modified");
	if (pCpCorePropertiesDctermsModified == nullptr)
	{
		return 1;
	}
	const char* pCpCorePropertiesDctermsModifiedText = pCpCorePropertiesDctermsModified->GetText();
	if (pCpCorePropertiesDctermsModifiedText != nullptr)
	{
		m_sLastModifiedDateTime = pCpCorePropertiesDctermsModifiedText;
	}
	UString sSwitchGamesCoreXmlPath = m_sModuleDirName + USTR("/Switch Games/docProps/core.xml");
	fp = UFopen(sSwitchGamesCoreXmlPath.c_str(), USTR("rb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	eError = xmlDoc.LoadFile(fp);
	fclose(fp);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		return 1;
	}
	pDocCpCoreProperties = xmlDoc.RootElement();
	pCpCorePropertiesDcCreator = pDocCpCoreProperties->FirstChildElement("dc:creator");
	pCpCorePropertiesDcCreator->SetText(m_sCreator.c_str());
	pCpCorePropertiesCpLastModifiedBy = pDocCpCoreProperties->FirstChildElement("cp:lastModifiedBy");
	pCpCorePropertiesCpLastModifiedBy->SetText(m_sLastModifiedBy.c_str());
	pCpCorePropertiesDctermsModified = pDocCpCoreProperties->FirstChildElement("dcterms:modified");
	pCpCorePropertiesDctermsModified->SetText(m_sLastModifiedDateTime.c_str());
	fp = UFopen(sCoreXmlPath.c_str(), USTR("wb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	eError = xmlDoc.SaveFile(fp, m_bCompact);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		fclose(fp);
		return 1;
	}
	fclose(fp);
	return 0;
}

int CSwitchGamesXlsx::resaveWorkbookRels() const
{
	if (!m_bResave)
	{
		return 1;
	}
	UString sWorkbookXmlRelsPath = m_sXlsxDirName + USTR("/xl/_rels/workbook.xml.rels");
	tinyxml2::XMLDocument relsDoc;
	UString sSwitchGamesWorkbookXmlRelsPath = m_sModuleDirName + USTR("/Switch Games/xl/_rels/workbook.xml.rels");
	FILE* fp = UFopen(sSwitchGamesWorkbookXmlRelsPath.c_str(), USTR("rb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	tinyxml2::XMLError eError = relsDoc.LoadFile(fp);
	fclose(fp);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		return 1;
	}
	tinyxml2::XMLElement* pDocRelationships = relsDoc.RootElement();
	pDocRelationships->DeleteChildren();
	tinyxml2::XMLElement* pRelationshipsRelationship = pDocRelationships->InsertNewChildElement("Relationship");
	pRelationshipsRelationship->SetAttribute("Id", Format("rId%d", static_cast<n32>(m_vSheetName.size()) + 3).c_str());
	pRelationshipsRelationship->SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
	pRelationshipsRelationship->SetAttribute("Target", "sharedStrings.xml");
	pRelationshipsRelationship = pDocRelationships->InsertNewChildElement("Relationship");
	pRelationshipsRelationship->SetAttribute("Id", Format("rId%d", static_cast<n32>(m_vSheetName.size()) + 2).c_str());
	pRelationshipsRelationship->SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
	pRelationshipsRelationship->SetAttribute("Target", "styles.xml");
	pRelationshipsRelationship = pDocRelationships->InsertNewChildElement("Relationship");
	pRelationshipsRelationship->SetAttribute("Id", Format("rId%d", static_cast<n32>(m_vSheetName.size()) + 1).c_str());
	pRelationshipsRelationship->SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme");
	pRelationshipsRelationship->SetAttribute("Target", "theme/theme1.xml");
	for (n32 i = static_cast<n32>(m_vSheetName.size()); i > 0; i--)
	{
		pRelationshipsRelationship = pDocRelationships->InsertNewChildElement("Relationship");
		pRelationshipsRelationship->SetAttribute("Id", Format("rId%d", i).c_str());
		pRelationshipsRelationship->SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
		pRelationshipsRelationship->SetAttribute("Target", Format("worksheets/sheet%d.xml", i).c_str());
	}
	fp = UFopen(sWorkbookXmlRelsPath.c_str(), USTR("wb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	eError = relsDoc.SaveFile(fp, m_bCompact);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		fclose(fp);
		return 1;
	}
	fclose(fp);
	return 0;
}

int CSwitchGamesXlsx::resaveTheme1() const
{
	if (!m_bResave)
	{
		return 1;
	}
	UString sTheme1XmlPath = m_sXlsxDirName + USTR("/xl/theme/theme1.xml");
	UString sSwitchGamesTheme1XmlPath = m_sModuleDirName + USTR("/Switch Games/xl/theme/theme1.xml");
	if (!copyFile(sTheme1XmlPath, sSwitchGamesTheme1XmlPath))
	{
		return 1;
	}
	return 0;
}

int CSwitchGamesXlsx::readSheet()
{
	n32 nSheetStringCount = 0;
	for (n32 i = 0; i < static_cast<n32>(m_vSheetName.size()); i++)
	{
		const wstring& sTableName = m_vSheetName[i];
		SSheetInfo& sheetInfo = mSheetInfo[sTableName];
		sheetInfo.TabSelected = i == m_nActiveTabNew;
		map<n32, n32>& mRowStyle = m_mTableRowStyle[sTableName];
		n32 nSheetIndex = i + 1;
		if (m_bResave)
		{
			nSheetIndex = m_mSheetIndexOld[sTableName];
			if (nSheetIndex < 1)
			{
				sheetInfo.RowCount = 2;
				sheetInfo.ColumnCount = 4;
				sheetInfo.Width.resize(4, 9);
				continue;
			}
		}
		UString sSheetXmlPath = m_sXlsxDirName + Format(USTR("/xl/worksheets/sheet%d.xml"), nSheetIndex);
		FILE* fp = UFopen(sSheetXmlPath.c_str(), USTR("rb"), false);
		if (fp == nullptr)
		{
			return 1;
		}
		fseek(fp, 0, SEEK_END);
		u32 uSheetXmlSize = ftell(fp);
		fseek(fp, 0, SEEK_SET);
		char* pSheetXml0 = new char[uSheetXmlSize + 1];
		fread(pSheetXml0, 1, uSheetXmlSize, fp);
		pSheetXml0[uSheetXmlSize] = 0;
		string sSheetXml0 = pSheetXml0;
		delete[] pSheetXml0;
		string sSheetXml;
		sSheetXml.reserve(uSheetXmlSize * 2);
		fseek(fp, 0, SEEK_SET);
		tinyxml2::XMLDocument xmlDoc;
		tinyxml2::XMLError eError = xmlDoc.LoadFile(fp);
		fclose(fp);
		if (eError != tinyxml2::XML_SUCCESS)
		{
			return 1;
		}
		if (!StartWith(sSheetXml0, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"))
		{
			return 1;
		}
		sSheetXml += "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
		const tinyxml2::XMLElement* pDocWorksheet = xmlDoc.RootElement();
		if (pDocWorksheet == nullptr)
		{
			return 1;
		}
		string sWorksheetName = pDocWorksheet->Name();
		if (sWorksheetName != "worksheet")
		{
			return 1;
		}
		sSheetXml += "<";
		sSheetXml += pDocWorksheet->Name();
		for (const tinyxml2::XMLAttribute* pWorksheetAttribute = pDocWorksheet->FirstAttribute(); pWorksheetAttribute != nullptr; pWorksheetAttribute = pWorksheetAttribute->Next())
		{
			string sAttributeName = pWorksheetAttribute->Name();
			string sAttributeValue = validateValue(pWorksheetAttribute->Value());
			sSheetXml += " " + sAttributeName + "=\"" + sAttributeValue + "\"";
		}
		sSheetXml += ">";
		for (const tinyxml2::XMLNode* pWorksheetChild = pDocWorksheet->FirstChild(); pWorksheetChild != nullptr; pWorksheetChild = pWorksheetChild->NextSibling())
		{
			if (pWorksheetChild->ToText() != nullptr)
			{
				return 1;
			}
			const tinyxml2::XMLElement* pWorksheetChildElement = pWorksheetChild->ToElement();
			if (pWorksheetChildElement == nullptr)
			{
				return 1;
			}
			string sWorksheetChildElementName = pWorksheetChildElement->Name();
			if (sWorksheetChildElementName != "sheetData")
			{
				if (!travelElement(pWorksheetChildElement, sSheetXml))
				{
					return 1;
				}
				if (sWorksheetChildElementName == "sheetViews")
				{
					const tinyxml2::XMLElement* pWorksheetSheetViews = pWorksheetChildElement;
					const tinyxml2::XMLElement* pSheetViewsSheetView = pWorksheetSheetViews->FirstChildElement("sheetView");
					if (pSheetViewsSheetView != nullptr)
					{
						const tinyxml2::XMLElement* pSheetViewPane = pSheetViewsSheetView->FirstChildElement("pane");
						if (pSheetViewPane != nullptr)
						{
							const char* pPaneTopLeftCell = pSheetViewPane->Attribute("topLeftCell");
							if (pPaneTopLeftCell != nullptr)
							{
								string sPaneTopLeftCell = pPaneTopLeftCell;
								string::size_type uRowPos = sPaneTopLeftCell.find_first_of("0123456789");
								sheetInfo.TopLeftCellRowIndex = SToN32(sPaneTopLeftCell.c_str() + uRowPos) - 1;
								string sIndex26 = sPaneTopLeftCell.substr(0, uRowPos);
								sheetInfo.TopLeftCellColumnIndex = 0;
								for (n32 i = 0; i < static_cast<n32>(sIndex26.size()); i++)
								{
									if (sIndex26[i] < 'A' || sIndex26[i] > 'Z')
									{
										return 1;
									}
									sheetInfo.TopLeftCellColumnIndex = sheetInfo.TopLeftCellColumnIndex * 26 + sIndex26[i] - 'A' + 1;
								}
								sheetInfo.TopLeftCellColumnIndex--;
							}
						}
						const tinyxml2::XMLElement* pSheetViewSelection = pSheetViewsSheetView->FirstChildElement("selection");
						if (pSheetViewSelection != nullptr)
						{
							const char* pSelectionActiveCell = pSheetViewSelection->Attribute("activeCell");
							if (pSelectionActiveCell != nullptr)
							{
								string sSelectionActiveCell = pSelectionActiveCell;
								string::size_type uRowPos = sSelectionActiveCell.find_first_of("0123456789");
								sheetInfo.ActiveCellRowIndex = SToN32(sSelectionActiveCell.c_str() + uRowPos) - 1;
								string sIndex26 = sSelectionActiveCell.substr(0, uRowPos);
								sheetInfo.ActiveCellColumnIndex = 0;
								for (n32 i = 0; i < static_cast<n32>(sIndex26.size()); i++)
								{
									if (sIndex26[i] < 'A' || sIndex26[i] > 'Z')
									{
										return 1;
									}
									sheetInfo.ActiveCellColumnIndex = sheetInfo.ActiveCellColumnIndex * 26 + sIndex26[i] - 'A' + 1;
								}
								sheetInfo.ActiveCellColumnIndex--;
							}
						}
					}
				}
			}
			else
			{
				const tinyxml2::XMLElement* pWorksheetSheetData = pWorksheetChildElement;
				sSheetXml += "<";
				sSheetXml += pWorksheetSheetData->Name();
				for (const tinyxml2::XMLAttribute* pSheetDataAttribute = pWorksheetSheetData->FirstAttribute(); pSheetDataAttribute != nullptr; pSheetDataAttribute = pSheetDataAttribute->Next())
				{
					// no attribute
					return 1;
					string sAttributeName = pSheetDataAttribute->Name();
					string sAttributeValue = validateValue(pSheetDataAttribute->Value());
					sSheetXml += " " + sAttributeName + "=\"" + sAttributeValue + "\"";
				}
				sSheetXml += ">";
				for (const tinyxml2::XMLNode* pSheetDataChild = pWorksheetSheetData->FirstChild(); pSheetDataChild != nullptr; pSheetDataChild = pSheetDataChild->NextSibling())
				{
					if (pSheetDataChild->ToText() != nullptr)
					{
						return 1;
					}
					const tinyxml2::XMLElement* pSheetDataChildElement = pSheetDataChild->ToElement();
					if (pSheetDataChildElement == nullptr)
					{
						return 1;
					}
					string sSheetDataChildElementName = pSheetDataChildElement->Name();
					if (sSheetDataChildElementName != "row")
					{
						// no other child element
						return 1;
						if (!travelElement(pSheetDataChildElement, sSheetXml))
						{
							return 1;
						}
					}
					else
					{
						const tinyxml2::XMLElement* pSheetDataRow = pSheetDataChildElement;
						n32 nRowIndex = -1;
						n32 nRowStyleOld = -1;
						n32 nRowStyleNew = kStyleIdNone;
						sSheetXml += "<";
						sSheetXml += pSheetDataRow->Name();
						for (const tinyxml2::XMLAttribute* pRowAttribute = pSheetDataRow->FirstAttribute(); pRowAttribute != nullptr; pRowAttribute = pRowAttribute->Next())
						{
							string sAttributeName = pRowAttribute->Name();
							string sAttributeValue = pRowAttribute->Value();
							if (sAttributeName == "r")
							{
								nRowIndex = SToN32(sAttributeValue) - 1;
							}
							else if (sAttributeName == "s")
							{
								nRowStyleOld = SToN32(sAttributeValue);
								nRowStyleNew = nRowStyleOld;
								if (m_bResave)
								{
									nRowStyleNew = m_mStyleOldToNew[nRowStyleOld];
								}
							}
							sAttributeValue = validateValue(sAttributeValue);
							sSheetXml += " " + sAttributeName + "=\"" + sAttributeValue + "\"";
						}
						sSheetXml += ">";
						if (nRowIndex < 0)
						{
							return 1;
						}
						if (nRowIndex + 1 > sheetInfo.RowCount)
						{
							sheetInfo.RowCount = nRowIndex + 1;
						}
						for (const tinyxml2::XMLNode* pRowChild = pSheetDataRow->FirstChild(); pRowChild != nullptr; pRowChild = pRowChild->NextSibling())
						{
							if (pRowChild->ToText() != nullptr)
							{
								return 1;
							}
							const tinyxml2::XMLElement* pRowChildElement = pRowChild->ToElement();
							if (pRowChildElement == nullptr)
							{
								return 1;
							}
							string sRowChildElementName = pRowChildElement->Name();
							if (sRowChildElementName != "c")
							{
								// no other child element
								return 1;
								if (!travelElement(pRowChildElement, sSheetXml))
								{
									return 1;
								}
							}
							else
							{
								const tinyxml2::XMLElement* pRowC = pRowChildElement;
								n32 nColumnIndex = -1;
								n32 nColumnStyleOld = -1;
								bool bCTypeString = false;
								sSheetXml += "<";
								sSheetXml += pRowC->Name();
								for (const tinyxml2::XMLAttribute* pCAttribute = pRowC->FirstAttribute(); pCAttribute != nullptr; pCAttribute = pCAttribute->Next())
								{
									string sAttributeName = pCAttribute->Name();
									string sAttributeValue = pCAttribute->Value();
									if (sAttributeName == "r")
									{
										string sRowAttributeRValue = Format("%d", nRowIndex + 1);
										if (!EndWith(sAttributeValue, sRowAttributeRValue))
										{
											return 1;
										}
										string sIndex26 = sAttributeValue.substr(0, sAttributeValue.size() - sRowAttributeRValue.size());
										nColumnIndex = 0;
										for (n32 i = 0; i < static_cast<n32>(sIndex26.size()); i++)
										{
											if (sIndex26[i] < 'A' || sIndex26[i] > 'Z')
											{
												return 1;
											}
											nColumnIndex = nColumnIndex * 26 + sIndex26[i] - 'A' + 1;
										}
										nColumnIndex--;
									}
									else if (sAttributeName == "s")
									{
										nColumnStyleOld = SToN32(sAttributeValue);
										if (nRowStyleOld == -1)
										{
											nRowStyleOld = nColumnStyleOld;
											nRowStyleNew = nRowStyleOld;
											if (m_bResave)
											{
												nRowStyleNew = m_mStyleOldToNew[nRowStyleOld];
											}
										}
										else if (nColumnStyleOld != nRowStyleOld)
										{
											nRowStyleNew = kStyleIdRed;
										}
									}
									else if (sAttributeName == "t")
									{
										if (sAttributeValue == "s")
										{
											bCTypeString = true;
										}
										else if (sAttributeValue == "str")
										{
											// do nothing
										}
										else
										{
											return 1;
										}
									}
									sAttributeValue = validateValue(sAttributeValue);
									sSheetXml += " " + sAttributeName + "=\"" + sAttributeValue + "\"";
								}
								if (nColumnIndex < 0)
								{
									return 1;
								}
								if (nColumnIndex + 1 > sheetInfo.ColumnCount)
								{
									sheetInfo.ColumnCount = nColumnIndex + 1;
									sheetInfo.Width.resize(sheetInfo.ColumnCount, 9);
								}
								const tinyxml2::XMLNode* pCChild = pRowC->FirstChild();
								if (pCChild == nullptr)
								{
									sSheetXml += "/>";
									mTableRowColumnText[sTableName][nRowIndex][nColumnIndex] = make_pair(false, L"");
								}
								else
								{
									sSheetXml += ">";
									for (; pCChild != nullptr; pCChild = pCChild->NextSibling())
									{
										if (pCChild->ToText() != nullptr)
										{
											return 1;
										}
										const tinyxml2::XMLElement* pCChildElement = pCChild->ToElement();
										if (pCChildElement == nullptr)
										{
											return 1;
										}
										string sCChildElementName = pCChildElement->Name();
										if (sCChildElementName != "f" && sCChildElementName != "v")
										{
											// no other child element
											return 1;
											if (!travelElement(pCChildElement, sSheetXml))
											{
												return 1;
											}
										}
										else if (sCChildElementName == "f")
										{
											if (!travelElement(pCChildElement, sSheetXml))
											{
												return 1;
											}
										}
										else if (sCChildElementName == "v")
										{
											const tinyxml2::XMLElement* pCV = pCChildElement;
											string sStmt;
											sSheetXml += "<";
											sSheetXml += pCV->Name();
											for (const tinyxml2::XMLAttribute* pVAttribute = pCV->FirstAttribute(); pVAttribute != nullptr; pVAttribute = pVAttribute->Next())
											{
												// no attribute
												return 1;
												string sAttributeName = pVAttribute->Name();
												string sAttributeValue = validateValue(pVAttribute->Value());
												sSheetXml += " " + sAttributeName + "=\"" + sAttributeValue + "\"";
											}
											sSheetXml += ">";
											n32 nVChildCount = 0;
											for (const tinyxml2::XMLNode* pVChild = pCV->FirstChild(); pVChild != nullptr; pVChild = pVChild->NextSibling())
											{
												if (nVChildCount != 0)
												{
													return 1;
												}
												const tinyxml2::XMLText* pVChildText = pVChild->ToText();
												if (pVChildText == nullptr)
												{
													return 1;
												}
												sStmt = pVChildText->Value();
												if (bCTypeString)
												{
													n32 nStringIndex = SToN32(sStmt);
													map<n32, wstring>::const_iterator itSharedStrings = m_mSharedStrings.find(nStringIndex);
													if (itSharedStrings == m_mSharedStrings.end())
													{
														return 1;
													}
													sStmt = WToU8(itSharedStrings->second);
													nSheetStringCount++;
												}
												sSheetXml += validateText(pVChildText->Value());
												nVChildCount++;
											}
											sSheetXml += "</";
											sSheetXml += pCV->Name();
											sSheetXml += ">";
											if (nVChildCount != 1)
											{
												return 1;
											}
											if (sStmt.empty())
											{
												return 1;
											}
											wstring sStmtW = U8ToW(sStmt);
											if (static_cast<n32>(sStmtW.size()) > sheetInfo.Width[nColumnIndex])
											{
												sheetInfo.Width[nColumnIndex] = static_cast<n32>(sStmtW.size());
											}
											mTableRowColumnText[sTableName][nRowIndex][nColumnIndex] = make_pair(true, sStmtW);
										}
									}
									sSheetXml += "</";
									sSheetXml += pRowC->Name();
									sSheetXml += ">";
								}
							}
						}
						mRowStyle[nRowIndex] = nRowStyleNew;
						sSheetXml += "</";
						sSheetXml += pSheetDataRow->Name();
						sSheetXml += ">";
					}
				}
				sSheetXml += "</";
				sSheetXml += pWorksheetSheetData->Name();
				sSheetXml += ">";
			}
		}
		sSheetXml += "</";
		sSheetXml += pDocWorksheet->Name();
		sSheetXml += ">";
		if (sSheetXml != sSheetXml0)
		{
			return 1;
		}
	}
	if (nSheetStringCount > m_nSstCount)
	{
		return 1;
	}
	if (m_bResave)
	{
		for (vector<wstring>::const_iterator it = m_vSheetName.begin(); it != m_vSheetName.end(); ++it)
		{
			const wstring& sTableName = *it;
			SSheetInfo& sheetInfo = mSheetInfo[sTableName];
			map<n32, n32>& mRowStyle = m_mTableRowStyle[sTableName];
			mRowStyle[0] = kStyleIdBlueBold;
			mRowStyle[1] = kStyleIdBlue;
			map<n32, map<n32, pair<bool, wstring>>>& mRowColumnText = mTableRowColumnText[sTableName];
			mRowColumnText[0][0] = make_pair(true, L"Release Name");
			mRowColumnText[0][1] = make_pair(true, L"File Name");
			mRowColumnText[0][2] = make_pair(true, L"Comment");
			mRowColumnText[0][3] = make_pair(true, L"Comment2");
			mRowColumnText[1][0] = make_pair(true, L"..");
			mRowColumnText[1][1] = make_pair(true, L"..");
			mRowColumnText[1][2] = make_pair(true, L"nfo LF/sfv CRLF");
			mRowColumnText[1].erase(3);
			sheetInfo.RowCount = max<n32>(2, sheetInfo.RowCount);
			sheetInfo.ColumnCount = max<n32>(4, sheetInfo.ColumnCount);
		}
	}
	return 0;
}

int CSwitchGamesXlsx::writeSheet()
{
	m_nSstCount = 0;
	for (n32 i = 0; i < static_cast<n32>(m_vSheetName.size()); i++)
	{
		const wstring& sTableName = m_vSheetName[i];
		SSheetInfo& sheetInfo = mSheetInfo[sTableName];
		map<n32, n32>& mRowStyle = m_mTableRowStyle[sTableName];
		map<n32, map<n32, pair<bool, wstring>>>& mRowColumnText = mTableRowColumnText[sTableName];
		string sSheetXml;
		sSheetXml += "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
		sSheetXml += "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac xr xr2 xr3\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\" xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\" xr:uid=\"{00000000-0001-0000-0000-000000000000}\">";
		sSheetXml += "<dimension ref=\"A1";
		if (sheetInfo.RowCount != 0 || sheetInfo.ColumnCount != 0)
		{
			sSheetXml += ":";
			sSheetXml += encodePosition(max<n32>(0, sheetInfo.RowCount - 1), max<n32>(0, sheetInfo.ColumnCount - 1));
		}
		sSheetXml += "\"/>";
		sSheetXml += "<sheetViews>";
		sSheetXml += "<sheetView";
		if (sheetInfo.TabSelected)
		{
			sSheetXml += " tabSelected=\"1\"";
		}
		sSheetXml += " workbookViewId=\"0\">";
		sSheetXml += "<pane ySplit=\"2\"";
		if (sheetInfo.TopLeftCellRowIndex != 0 || sheetInfo.TopLeftCellColumnIndex != 0)
		{
			sSheetXml += " topLeftCell=\"";
			sSheetXml += encodePosition(sheetInfo.TopLeftCellRowIndex, sheetInfo.TopLeftCellColumnIndex);
			sSheetXml += "\"";
		}
		sSheetXml += " activePane=\"bottomLeft\" state=\"frozen\"/>";
		sSheetXml += "<selection pane=\"bottomLeft\"";
		if (sheetInfo.ActiveCellRowIndex != 0 || sheetInfo.ActiveCellColumnIndex != 0)
		{
			string sActiveCell = encodePosition(sheetInfo.ActiveCellRowIndex, sheetInfo.ActiveCellColumnIndex);
			sSheetXml += " activeCell=\"" + sActiveCell + "\"";
			sSheetXml += " sqref=\"" + sActiveCell + "\"";
		}
		sSheetXml += "/>";
		sSheetXml += "</sheetView>";
		sSheetXml += "</sheetViews>";
		sSheetXml += "<sheetFormatPr defaultRowHeight=\"14.25\" x14ac:dyDescent=\"0.2\"/>";
		sSheetXml += "<cols>";
		for (n32 j = 0; j < sheetInfo.ColumnCount; j++)
		{
			sSheetXml += Format("<col min=\"%d\" max=\"%d\" width=\"%d\" style=\"%d\" bestFit=\"1\" customWidth=\"1\"/>", j + 1, j + 1, sheetInfo.Width[j], kStyleIdNone);
		}
		sSheetXml += Format("<col min=\"%d\" max=\"16384\" width=\"9\" style=\"%d\"/>", sheetInfo.ColumnCount + 1, kStyleIdNone);
		sSheetXml += "</cols>";
		sSheetXml += "<sheetData>";
		for (map<n32, map<n32, pair<bool, wstring>>>::const_iterator itRow = mRowColumnText.begin(); itRow != mRowColumnText.end(); ++itRow)
		{
			n32 nRowIndex = itRow->first;
			const map<n32, pair<bool, wstring>>& mColumnText = itRow->second;
			n32 nRowStyle = mRowStyle[nRowIndex];
			sSheetXml += Format("<row r=\"%d\" spans=\"1:%d\" s=\"%d\" customFormat=\"1\" x14ac:dyDescent=\"0.2\">", nRowIndex + 1, sheetInfo.ColumnCount, nRowStyle);
			for (map<n32, pair<bool, wstring>>::const_iterator itColumn = mColumnText.begin(); itColumn != mColumnText.end(); ++itColumn)
			{
				n32 nColumnIndex = itColumn->first;
				const pair<bool, wstring>& pText = itColumn->second;
				if (!pText.first)
				{
					continue;
				}
				m_nSstCount++;
				const wstring& sText = pText.second;
				sSheetXml += Format("<c r=\"%s\" s=\"%d\" t=\"s\">", encodePosition(nRowIndex, nColumnIndex).c_str(), nRowStyle);
				sSheetXml += Format("<v>%d</v>", m_mSharedStringsIndexNew[sText]);
				sSheetXml += "</c>";
			}
			sSheetXml += "</row>";
		}
		sSheetXml += "</sheetData>";
		sSheetXml += "<phoneticPr fontId=\"1\" type=\"noConversion\"/>";
		sSheetXml += "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>";
		sSheetXml += "<pageSetup paperSize=\"9\" orientation=\"portrait\" r:id=\"rId1\"/>";
		sSheetXml += "</worksheet>";
		UString sSheetXmlPath = m_sXlsxDirName + Format(USTR("/xl/worksheets/sheet%d.xml"), i + 1);
		FILE* fp = UFopen(sSheetXmlPath.c_str(), USTR("wb"), false);
		if (fp == nullptr)
		{
			return 1;
		}
		fwrite(sSheetXml.c_str(), 1, sSheetXml.size(), fp);
		fclose(fp);
	}
	return 0;
}

int CSwitchGamesXlsx::writeSharedStrings() const
{
	UString sSharedStringsXmlPath = m_sXlsxDirName + USTR("/xl/sharedStrings.xml");
	string sSharedStringsXml;
	sSharedStringsXml += "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
	sSharedStringsXml += Format("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"%d\" uniqueCount=\"%d\">", m_nSstCount, static_cast<n32>(m_mSharedStringsIndexNew.size()));
	for (map<wstring, n32>::const_iterator it = m_mSharedStringsIndexNew.begin(); it != m_mSharedStringsIndexNew.end(); ++it)
	{
		string sTextU8 = WToU8(it->first);
		sSharedStringsXml += "<si>";
		sSharedStringsXml += "<t>";
		sSharedStringsXml += sTextU8;
		sSharedStringsXml += "</t>";
		sSharedStringsXml += "<phoneticPr fontId=\"1\" type=\"noConversion\"/>";
		sSharedStringsXml += "</si>";
	}
	sSharedStringsXml += "</sst>";
	FILE* fp = UFopen(sSharedStringsXmlPath.c_str(), USTR("wb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	fwrite(sSharedStringsXml.c_str(), 1, sSharedStringsXml.size(), fp);
	fclose(fp);
	return 0;
}

int CSwitchGamesXlsx::resaveStyles() const
{
	if (!m_bResave)
	{
		return 1;
	}
	UString sStylesXmlPath = m_sXlsxDirName + USTR("/xl/styles.xml");
	UString sSwitchGamesStylesXmlPath = m_sModuleDirName + USTR("/Switch Games/xl/styles.xml");
	if (!copyFile(sStylesXmlPath, sSwitchGamesStylesXmlPath))
	{
		return 1;
	}
	return 0;
}

int CSwitchGamesXlsx::resaveWorkbook() const
{
	if (!m_bResave)
	{
		return 1;
	}
	UString sWorkbookXmlPath = m_sXlsxDirName + USTR("/xl/workbook.xml");
	tinyxml2::XMLDocument xmlDoc;
	UString sSwitchGamesWorkbookXmlPath = m_sModuleDirName + USTR("/Switch Games/xl/workbook.xml");
	FILE* fp = UFopen(sSwitchGamesWorkbookXmlPath.c_str(), USTR("rb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	tinyxml2::XMLError eError = xmlDoc.LoadFile(fp);
	fclose(fp);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		return 1;
	}
	tinyxml2::XMLElement* pDocWorkbook = xmlDoc.RootElement();
	tinyxml2::XMLElement* pWorkbookBookViews = pDocWorkbook->FirstChildElement("bookViews");
	if (pWorkbookBookViews != nullptr)
	{
		tinyxml2::XMLElement* pBookViewsWorkbookView = pWorkbookBookViews->FirstChildElement("workbookView");
		if (pBookViewsWorkbookView != nullptr)
		{
			if (m_nActiveTabNew == 0)
			{
				pBookViewsWorkbookView->DeleteAttribute("activeTab");
			}
			else
			{
				pBookViewsWorkbookView->SetAttribute("activeTab", m_nActiveTabNew);
			}
		}
	}
	tinyxml2::XMLElement* pWorkbookSheets = pDocWorkbook->FirstChildElement("sheets");
	pWorkbookSheets->DeleteChildren();
	for (n32 i = 0; i < static_cast<n32>(m_vSheetName.size()); i++)
	{
		tinyxml2::XMLElement* pSheetsSheet = pWorkbookSheets->InsertNewChildElement("sheet");
		pSheetsSheet->SetAttribute("name", WToU8(m_vSheetName[i]).c_str());
		pSheetsSheet->SetAttribute("sheetId", i + 1);
		pSheetsSheet->SetAttribute("r:id", Format("rId%d", i + 1).c_str());
	}
	fp = UFopen(sWorkbookXmlPath.c_str(), USTR("wb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	eError = xmlDoc.SaveFile(fp, m_bCompact);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		fclose(fp);
		return 1;
	}
	fclose(fp);
	return 0;
}

int CSwitchGamesXlsx::resaveContentTypes() const
{
	if (!m_bResave)
	{
		return 1;
	}
	UString sContentTypesXmlPath = m_sXlsxDirName + USTR("/[Content_Types].xml");
	tinyxml2::XMLDocument xmlDoc;
	UString sSwitchGamesContentTypesXmlPath = m_sModuleDirName + USTR("/Switch Games/[Content_Types].xml");
	FILE* fp = UFopen(sSwitchGamesContentTypesXmlPath.c_str(), USTR("rb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	tinyxml2::XMLError eError = xmlDoc.LoadFile(fp);
	fclose(fp);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		return 1;
	}
	tinyxml2::XMLElement* pDocTypes = xmlDoc.RootElement();
	pDocTypes->DeleteChildren();
	tinyxml2::XMLElement* pTypesDefault = pDocTypes->InsertNewChildElement("Default");
	pTypesDefault->SetAttribute("Extension", "rels");
	pTypesDefault->SetAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
	pTypesDefault = pDocTypes->InsertNewChildElement("Default");
	pTypesDefault->SetAttribute("Extension", "xml");
	pTypesDefault->SetAttribute("ContentType", "application/xml");
	tinyxml2::XMLElement* pTypesOverride = pDocTypes->InsertNewChildElement("Override");
	pTypesOverride->SetAttribute("PartName", "/xl/workbook.xml");
	pTypesOverride->SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
	for (n32 i = 0; i < static_cast<n32>(m_vSheetName.size()); i++)
	{
		pTypesOverride = pDocTypes->InsertNewChildElement("Override");
		pTypesOverride->SetAttribute("PartName", Format("/xl/worksheets/sheet%d.xml", i + 1).c_str());
		pTypesOverride->SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
	}
	pTypesOverride = pDocTypes->InsertNewChildElement("Override");
	pTypesOverride->SetAttribute("PartName", "/xl/theme/theme1.xml");
	pTypesOverride->SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.theme+xml");
	pTypesOverride = pDocTypes->InsertNewChildElement("Override");
	pTypesOverride->SetAttribute("PartName", "/xl/styles.xml");
	pTypesOverride->SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
	pTypesOverride = pDocTypes->InsertNewChildElement("Override");
	pTypesOverride->SetAttribute("PartName", "/xl/sharedStrings.xml");
	pTypesOverride->SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
	pTypesOverride = pDocTypes->InsertNewChildElement("Override");
	pTypesOverride->SetAttribute("PartName", "/docProps/core.xml");
	pTypesOverride->SetAttribute("ContentType", "application/vnd.openxmlformats-package.core-properties+xml");
	pTypesOverride = pDocTypes->InsertNewChildElement("Override");
	pTypesOverride->SetAttribute("PartName", "/docProps/app.xml");
	pTypesOverride->SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
	fp = UFopen(sContentTypesXmlPath.c_str(), USTR("wb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	eError = xmlDoc.SaveFile(fp, m_bCompact);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		fclose(fp);
		return 1;
	}
	fclose(fp);
	return 0;
}

int CSwitchGamesXlsx::readTable()
{
	for (n32 i = 0; i < static_cast<n32>(m_vSheetName.size()); i++)
	{
		const wstring& sTableName = m_vSheetName[i];
		SSheetInfo& sheetInfo = mSheetInfo[sTableName];
		map<n32, n32>& mRowStyle = m_mTableRowStyle[sTableName];
		map<n32, map<n32, pair<bool, wstring>>>& mRowColumnText = mTableRowColumnText[sTableName];
		UString sTableFileName = m_sTableDirName + USTR("/") + WToU(sTableName) + USTR(".tsv");
		FILE* fp = UFopen(sTableFileName.c_str(), USTR("rb"), false);
		if (fp == nullptr)
		{
			return 1;
		}
		fseek(fp, 0, SEEK_END);
		u32 uTableFileSize = ftell(fp);
		fseek(fp, 0, SEEK_SET);
		char* pTemp = new char[uTableFileSize + 1];
		fread(pTemp, 1, uTableFileSize, fp);
		fclose(fp);
		pTemp[uTableFileSize] = 0;
		string sTable = pTemp;
		delete[] pTemp;
		vector<string> vTable = SplitOf(sTable, "\r\n");
		for (vector<string>::iterator it = vTable.begin(); it != vTable.end(); ++it)
		{
			*it = trim(*it);
		}
		vector<string>::const_iterator itTable = remove_if(vTable.begin(), vTable.end(), empty);
		vTable.erase(itTable, vTable.end());
		bool bReadTable = false;
		n32 nRowIndex = 0;
		for (vector<string>::const_iterator it = vTable.begin(); it != vTable.end(); ++it)
		{
			const string& sLine = *it;
			vector<string> vLine = Split(sLine, "\t");
			if (vLine.size() < 2)
			{
				if (vLine.size() != 1 || vLine[0] != "[style]")
				{
					return 1;
				}
			}
			if (!bReadTable)
			{
				if (StartWith(vLine[0], "[RowCount]"))
				{
					sheetInfo.RowCount = SToN32(vLine[1]);
				}
				else if (StartWith(vLine[0], "[ColumnCount]"))
				{
					sheetInfo.ColumnCount = SToN32(vLine[1]);
				}
				else if (StartWith(vLine[0], "[TabSelected]"))
				{
					sheetInfo.TabSelected = SToN32(vLine[1]) != 0;
					if (sheetInfo.TabSelected)
					{
						m_nActiveTabNew = i;
					}
				}
				else if (StartWith(vLine[0], "[TopLeftCellRowIndex]"))
				{
					sheetInfo.TopLeftCellRowIndex = SToN32(vLine[1]);
				}
				else if (StartWith(vLine[0], "[TopLeftCellColumnIndex]"))
				{
					sheetInfo.TopLeftCellColumnIndex = SToN32(vLine[1]);
				}
				else if (StartWith(vLine[0], "[ActiveCellRowIndex]"))
				{
					sheetInfo.ActiveCellRowIndex = SToN32(vLine[1]);
				}
				else if (StartWith(vLine[0], "[ActiveCellColumnIndex]"))
				{
					sheetInfo.ActiveCellColumnIndex = SToN32(vLine[1]);
				}
				else if (StartWith(vLine[0], "[Width]"))
				{
					for (n32 j = 0; j < static_cast<n32>(vLine.size() - 1); j++)
					{
						sheetInfo.Width.push_back(SToN32(vLine[j + 1]));
					}
				}
				else if (StartWith(vLine[0], "[style]"))
				{
					if (sheetInfo.ColumnCount != sheetInfo.Width.size())
					{
						return 1;
					}
					bReadTable = true;
				}
				else if (StartWith(vLine[0], "["))
				{
					return 1;
				}
			}
			else
			{
				if (vLine[0] == "nb")
				{
					mRowStyle[nRowIndex] = kStyleIdNoneNoBorder;
				}
				else if (vLine[0] == "no")
				{
					mRowStyle[nRowIndex] = kStyleIdNone;
				}
				else if (vLine[0] == "re")
				{
					mRowStyle[nRowIndex] = kStyleIdRed;
				}
				else if (vLine[0] == "bb")
				{
					mRowStyle[nRowIndex] = kStyleIdBlueBold;
				}
				else if (vLine[0] == "bl")
				{
					mRowStyle[nRowIndex] = kStyleIdBlue;
				}
				else if (vLine[0] == "gr")
				{
					mRowStyle[nRowIndex] = kStyleIdGreen;
				}
				else if (vLine[0] == "ye")
				{
					mRowStyle[nRowIndex] = kStyleIdYellow;
				}
				else if (vLine[0] == "go")
				{
					mRowStyle[nRowIndex] = kStyleIdGold;
				}
				map<n32, pair<bool, wstring>>& mColumnText = mRowColumnText[nRowIndex];
				for (n32 nColumnIndex = 0; nColumnIndex < static_cast<n32>(vLine.size() - 1); nColumnIndex++)
				{
					if (vLine[nColumnIndex + 1].empty())
					{
						mColumnText[nColumnIndex] = make_pair(false, L"");
					}
					else
					{
						mColumnText[nColumnIndex] = make_pair(true, U8ToW(vLine[nColumnIndex + 1]));
					}
				}
				nRowIndex++;
			}
		}
	}
	return 0;
}

int CSwitchGamesXlsx::writeTable()
{
	if (!makeDir(m_sTableDirName.c_str()))
	{
		return 1;
	}
	for (n32 i = 0; i < static_cast<n32>(m_vSheetName.size()); i++)
	{
		const wstring& sTableName = m_vSheetName[i];
		SSheetInfo& sheetInfo = mSheetInfo[sTableName];
		map<n32, n32>& mRowStyle = m_mTableRowStyle[sTableName];
		map<n32, map<n32, pair<bool, wstring>>>& mRowColumnText = mTableRowColumnText[sTableName];
		UString sTableFileName = m_sTableDirName + USTR("/") + WToU(sTableName) + USTR(".tsv");
		FILE* fp = UFopen(sTableFileName.c_str(), USTR("wb"), false);
		if (fp == nullptr)
		{
			return 1;
		}
		fprintf(fp, "[RowCount]\t%d\r\n", sheetInfo.RowCount);
		fprintf(fp, "[ColumnCount]\t%d\r\n", sheetInfo.ColumnCount);
		fprintf(fp, "[TabSelected]\t%d\r\n", sheetInfo.TabSelected ? 1 : 0);
		fprintf(fp, "[TopLeftCellRowIndex]\t%d\r\n", sheetInfo.TopLeftCellRowIndex);
		fprintf(fp, "[TopLeftCellColumnIndex]\t%d\r\n", sheetInfo.TopLeftCellColumnIndex);
		fprintf(fp, "[ActiveCellRowIndex]\t%d\r\n", sheetInfo.ActiveCellRowIndex);
		fprintf(fp, "[ActiveCellColumnIndex]\t%d\r\n", sheetInfo.ActiveCellColumnIndex);
		fprintf(fp, "[Width]");
		for (vector<n32>::const_iterator it = sheetInfo.Width.begin(); it != sheetInfo.Width.end(); ++it)
		{
			n32 nWidth = *it;
			fprintf(fp, "\t%d", nWidth);
		}
		fprintf(fp, "\r\n");
		fprintf(fp, "\r\n");
		fprintf(fp, "[style]\r\n");
		for (n32 nRowIndex = 0; nRowIndex < sheetInfo.RowCount; nRowIndex++)
		{
			map<n32, pair<bool, wstring>>& mColumnText = mRowColumnText[nRowIndex];
			n32 nRowStyle = mRowStyle[nRowIndex];
			switch (nRowStyle)
			{
			case kStyleIdNoneNoBorder:
				fprintf(fp, "nb");
				break;
			case kStyleIdNone:
				fprintf(fp, "no");
				break;
			case kStyleIdRed:
				fprintf(fp, "re");
				break;
			case kStyleIdBlueBold:
				fprintf(fp, "bb");
				break;
			case kStyleIdBlue:
				fprintf(fp, "bl");
				break;
			case kStyleIdGreen:
				fprintf(fp, "gr");
				break;
			case kStyleIdYellow:
				fprintf(fp, "ye");
				break;
			case kStyleIdGold:
				fprintf(fp, "go");
				break;
			default:
				fprintf(fp, "re");
				break;
			}
			for (n32 nColumnIndex = 0; nColumnIndex < sheetInfo.ColumnCount; nColumnIndex++)
			{
				pair<bool, wstring>& pText = mColumnText[nColumnIndex];
				string sTextU8;
				if (pText.first)
				{
					sTextU8 = WToU8(pText.second);
				}
				fprintf(fp, "\t%s", sTextU8.c_str());
			}
			fprintf(fp, "\r\n");
		}
		fclose(fp);
	}
	return 0;
}

int CSwitchGamesXlsx::sortTable()
{
	for (n32 i = 0; i < static_cast<n32>(m_vSheetName.size()); i++)
	{
		const wstring& sTableName = m_vSheetName[i];
		SSheetInfo& sheetInfo = mSheetInfo[sTableName];
		map<n32, n32>& mRowStyle = m_mTableRowStyle[sTableName];
		map<n32, map<n32, pair<bool, wstring>>>& mRowColumnText = mTableRowColumnText[sTableName];
		vector<pair<n32, wstring>> vOrder;
		for (map<n32, map<n32, pair<bool, wstring>>>::iterator itRow = mRowColumnText.begin(); itRow != mRowColumnText.end(); ++itRow)
		{
			n32 nRowIndex = itRow->first;
			map<n32, pair<bool, wstring>>& mColumnText = mRowColumnText[nRowIndex];
			wstring sColumnTextUpper = mColumnText[0].second;
			transform(sColumnTextUpper.begin(), sColumnTextUpper.end(), sColumnTextUpper.begin(), ::towupper);
			vOrder.push_back(make_pair(nRowIndex, sColumnTextUpper));
		}
		stable_sort(vOrder.begin() + 2, vOrder.end(), rowColumnTextCompare);
		map<n32, n32> mRowIndexOldToNew;
		for (n32 nRowIndexNew = 0; nRowIndexNew < static_cast<n32>(vOrder.size()); nRowIndexNew++)
		{
			n32 nRowIndexOld = vOrder[nRowIndexNew].first;
			mRowIndexOldToNew[nRowIndexOld] = nRowIndexNew;
		}
		map<n32, n32>::iterator itRowIndex = mRowIndexOldToNew.find(sheetInfo.TopLeftCellRowIndex);
		if (itRowIndex != mRowIndexOldToNew.end())
		{
			sheetInfo.TopLeftCellRowIndex = itRowIndex->second;
		}
		itRowIndex = mRowIndexOldToNew.find(sheetInfo.ActiveCellRowIndex);
		if (itRowIndex != mRowIndexOldToNew.end())
		{
			sheetInfo.ActiveCellRowIndex = itRowIndex->second;
		}
		map<n32, n32> mRowStyleTemp;
		mRowStyleTemp.swap(mRowStyle);
		map<n32, map<n32, pair<bool, wstring>>> mRowColumnTextTemp;
		mRowColumnTextTemp.swap(mRowColumnText);
		for (map<n32, n32>::const_iterator itRowIndex = mRowIndexOldToNew.begin(); itRowIndex != mRowIndexOldToNew.end(); ++itRowIndex)
		{
			n32 nRowIndexOld = itRowIndex->first;
			n32 nRowIndexNew = itRowIndex->second;
			mRowStyle[nRowIndexNew] = mRowStyleTemp[nRowIndexOld];
			mRowColumnText[nRowIndexNew].swap(mRowColumnTextTemp[nRowIndexOld]);
		}
	}
	return 0;
}

int CSwitchGamesXlsx::checkTable()
{
	bool bAllExist = true;
	n32 nCheckIndex = 1;
	for (vector<SResult>::iterator itResult = m_vResult.begin(); itResult != m_vResult.end(); ++itResult)
	{
		SResult& result = *itResult;
		const wstring& sName = result.Name;
		UPrintf(USTR("\t%5d/%5d\t%") PRIUS USTR(" %4d %") PRIUS USTR("\n"), nCheckIndex, static_cast<n32>(m_vResult.size()), (result.Exist ? USTR("[v]     ") : USTR("    [x] ")), result.Year, WToU(sName).c_str());
		nCheckIndex++;
		if (!result.Exist)
		{
			bAllExist = false;
			continue;
		}
		UString sDirPath = WToU(result.Path);
		const wstring& sType = result.Type;
		map<n32, n32>& mRowStyle = m_mTableRowStyle[sType];
		map<n32, map<n32, pair<bool, wstring>>>& mRowColumnText = mTableRowColumnText[sType];
		n32 nRowIndex = -1;
		for (map<n32, map<n32, pair<bool, wstring>>>::iterator itRow = mRowColumnText.begin(); itRow != mRowColumnText.end(); ++itRow)
		{
			map<n32, pair<bool, wstring>>& mColumnText = itRow->second;
			if (mColumnText[0].first && mColumnText[0].second == sName)
			{
				nRowIndex = itRow->first;
				break;
			}
		}
		if (nRowIndex < 0)
		{
			UPrintf(USTR("NOT in table %") PRIUS USTR(": %") PRIUS USTR("\n"), WToU(sType).c_str(), sDirPath.c_str());
			continue;
		}
		if (matchGreenStyle(sName))
		{
			if (mRowStyle[nRowIndex] != kStyleIdGreen)
			{
				if (mRowStyle[nRowIndex] != kStyleIdRed)
				{
					mRowStyle[nRowIndex] = kStyleIdGreen;
				}
				else
				{
					mRowStyle[nRowIndex] = kStyleIdRed;
				}
			}
		}
		else
		{
			if (mRowStyle[nRowIndex] == kStyleIdGreen)
			{
				mRowStyle[nRowIndex] = kStyleIdRed;
			}
			if (bAllExist)
			{
				if (mRowStyle[nRowIndex] != kStyleIdYellow)
				{
					if (mRowStyle[nRowIndex] != kStyleIdRed)
					{
						mRowStyle[nRowIndex] = kStyleIdYellow;
					}
					else
					{
						mRowStyle[nRowIndex] = kStyleIdRed;
					}
				}
			}
			else
			{
				if (mRowStyle[nRowIndex] == kStyleIdYellow)
				{
					mRowStyle[nRowIndex] = kStyleIdRed;
				}
			}
		}
		UString::size_type uPrefixSize = sDirPath.size() + 1;
		vector<UString> vFile;
		queue<UString> qDir;
		qDir.push(sDirPath);
		while (!qDir.empty())
		{
			UString& sParent = qDir.front();
#if SDW_PLATFORM == SDW_PLATFORM_WINDOWS
			WIN32_FIND_DATAW ffd;
			HANDLE hFind = INVALID_HANDLE_VALUE;
			wstring sPattern = sParent + L"/*";
			hFind = FindFirstFileW(sPattern.c_str(), &ffd);
			if (hFind != INVALID_HANDLE_VALUE)
			{
				do
				{
					if ((ffd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
					{
						wstring sFileName = sParent + L"/" + ffd.cFileName;
						vFile.push_back(sFileName.substr(uPrefixSize));
					}
					else if ((ffd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0 && wcscmp(ffd.cFileName, L".") != 0 && wcscmp(ffd.cFileName, L"..") != 0)
					{
						wstring sDir = sParent + L"/" + ffd.cFileName;
						qDir.push(sDir);
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
					if (pDirent->d_type == DT_REG)
					{
						string sFileName = sParent + "/" + sName;
						vFile.push_back(sFileName.substr(uPrefixSize));
					}
					else if (pDirent->d_type == DT_DIR && strcmp(pDirent->d_name, ".") != 0 && strcmp(pDirent->d_name, "..") != 0)
					{
						string sDir = sParent + "/" + sName;
						qDir.push(sDir);
					}
				}
				closedir(pDir);
			}
#endif
			qDir.pop();
		}
		sort(vFile.begin(), vFile.end(), pathCompare);
		for (vector<UString>::const_iterator itFile = vFile.begin(); itFile != vFile.end(); ++itFile)
		{
			const UString& sFile = *itFile;
			UString::size_type uPos = sFile.rfind(USTR('.'));
			if (uPos == UString::npos)
			{
				result.OtherFile.push_back(sFile);
				continue;
			}
			UString sExtName = sFile.substr(uPos + 1);
			if (sExtName == USTR("rar"))
			{
				result.RarFile[sFile] = 0;
			}
			if (sExtName == USTR("nfo"))
			{
				result.NfoFile.push_back(sFile);
			}
			else if (sExtName == USTR("sfv"))
			{
				result.SfvFile.push_back(sFile);
			}
			else
			{
				if (sExtName.size() == 3)
				{
					if (sExtName[0] >= USTR('r') && sExtName[0] < USTR('|') && sExtName[1] >= USTR('0') && sExtName[1] <= USTR('9') && sExtName[2] >= USTR('0') && sExtName[2] <= USTR('9'))
					{
						result.RarFile[sFile] = 0;
					}
				}
				else
				{
					result.OtherFile.push_back(sFile);
				}
			}
		}
		if (result.NfoFile.size() > 1)
		{
			UPrintf(USTR("nfo COUNT %d > 1: %") PRIUS USTR("\n"), static_cast<n32>(result.NfoFile.size()), sDirPath.c_str());
			for (n32 i = 0; i < static_cast<n32>(result.NfoFile.size()); i++)
			{
				UPrintf(USTR("nfo[%d]: %") PRIUS USTR("\n"), i, result.NfoFile[i].c_str());
			}
			mRowStyle[nRowIndex] = kStyleIdRed;
		}
		if (result.SfvFile.size() > 1)
		{
			UPrintf(USTR("sfv COUNT %d > 1: %") PRIUS USTR("\n"), static_cast<n32>(result.SfvFile.size()), sDirPath.c_str());
			for (n32 i = 0; i < static_cast<n32>(result.SfvFile.size()); i++)
			{
				UPrintf(USTR("sfv[%d]: %") PRIUS USTR("\n"), i, result.SfvFile[i].c_str());
			}
			mRowStyle[nRowIndex] = kStyleIdRed;
		}
		if (result.OtherFile.size() > 1)
		{
			UPrintf(USTR("other COUNT %d > 1: %") PRIUS USTR("\n"), static_cast<n32>(result.OtherFile.size()), sDirPath.c_str());
			for (n32 i = 0; i < static_cast<n32>(result.OtherFile.size()); i++)
			{
				UPrintf(USTR("other[%d]: %") PRIUS USTR("\n"), i, result.OtherFile[i].c_str());
			}
			mRowStyle[nRowIndex] = kStyleIdRed;
		}
		wstring sCommentOld = mRowColumnText[nRowIndex][2].second;
		wstring sCommentNew;
		UString sPatchDirName = m_sTableDirName + USTR("/") + WToU(sType + L"/" + sName);
		if (!result.NfoFile.empty())
		{
			UString sFilePath = sDirPath + USTR("/") + result.NfoFile[0];
			STextFileContent textFileContent;
			if (readTextFile(sFilePath, textFileContent) != 0)
			{
				UPrintf(USTR("read text file error: %") PRIUS USTR("\n"), sFilePath.c_str());
				mRowStyle[nRowIndex] = kStyleIdRed;
				continue;
			}
			if (textFileContent.LineTypeNew == kLineTypeLF)
			{
				sCommentNew += L"/nfo LF";
				if (textFileContent.EncodingNew == kEncodingUTF8withBOM)
				{
					sCommentNew += L" with BOM";
				}
			}
			else if (textFileContent.LineTypeNew == kLineTypeLF_CR)
			{
				sCommentNew += L"/nfo LF|CR";
				if (textFileContent.EncodingNew == kEncodingUTF8withBOM)
				{
					sCommentNew += L" with BOM";
				}
			}
			else if (textFileContent.LineTypeNew == kLineTypeCRLF)
			{
				sCommentNew += L"/nfo CRLF";
				if (textFileContent.EncodingNew == kEncodingUTF8withBOM)
				{
					sCommentNew += L" with BOM";
				}
			}
			else
			{
				sCommentNew += L"/nfo MIX";
				mRowStyle[nRowIndex] = kStyleIdRed;
			}
			if (textFileContent.TextNew != textFileContent.TextOld || textFileContent.EncodingNew != textFileContent.EncodingOld || textFileContent.LineTypeNew != textFileContent.LineTypeOld)
			{
				UString sPatchFileName = sPatchDirName + USTR("/") + result.NfoFile[0];
				if (writeFileString(sPatchFileName, textFileContent.TextNew) != 0)
				{
					return 1;
				}
			}
		}
		if (!result.SfvFile.empty())
		{
			UString sFilePath = sDirPath + USTR("/") + result.SfvFile[0];
			STextFileContent textFileContent;
			if (readTextFile(sFilePath, textFileContent) != 0)
			{
				UPrintf(USTR("read text file error: %") PRIUS USTR("\n"), sFilePath.c_str());
				mRowStyle[nRowIndex] = kStyleIdRed;
				continue;
			}
			if (textFileContent.LineTypeNew == kLineTypeLF)
			{
				sCommentNew += L"/sfv LF";
				if (textFileContent.EncodingNew == kEncodingUTF8withBOM)
				{
					sCommentNew += L" with BOM";
				}
			}
			else if (textFileContent.LineTypeNew == kLineTypeLF_CR)
			{
				sCommentNew += L"/sfv LF|CR";
				if (textFileContent.EncodingNew == kEncodingUTF8withBOM)
				{
					sCommentNew += L" with BOM";
				}
			}
			else if (textFileContent.LineTypeNew == kLineTypeCRLF)
			{
				sCommentNew += L"/sfv CRLF";
				if (textFileContent.EncodingNew == kEncodingUTF8withBOM)
				{
					sCommentNew += L" with BOM";
				}
			}
			else
			{
				sCommentNew += L"/sfv MIX";
				mRowStyle[nRowIndex] = kStyleIdRed;
			}
			if (textFileContent.TextNew != textFileContent.TextOld || textFileContent.EncodingNew != textFileContent.EncodingOld || textFileContent.LineTypeNew != textFileContent.LineTypeOld)
			{
				UString sPatchFileName = sPatchDirName + USTR("/") + result.SfvFile[0];
				if (writeFileString(sPatchFileName, textFileContent.TextNew) != 0)
				{
					return 1;
				}
			}
		}
		if (!sCommentNew.empty())
		{
			sCommentNew.erase(0, 1);
		}
		mRowColumnText[nRowIndex][2].second = sCommentNew;
	}
	return 0;
}

int CSwitchGamesXlsx::readResult()
{
	FILE* fp = UFopen(m_sResultFileName.c_str(), USTR("rb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	fseek(fp, 0, SEEK_END);
	u32 uResultFileSize = ftell(fp);
	fseek(fp, 0, SEEK_SET);
	char* pTemp = new char[uResultFileSize + 1];
	fread(pTemp, 1, uResultFileSize, fp);
	fclose(fp);
	pTemp[uResultFileSize] = 0;
	string sResultText = pTemp;
	delete[] pTemp;
	vector<string> vResultText = SplitOf(sResultText, "\r\n");
	for (vector<string>::iterator itResultText = vResultText.begin(); itResultText != vResultText.end(); ++itResultText)
	{
		*itResultText = trim(*itResultText);
	}
	vector<string>::const_iterator itResultText = remove_if(vResultText.begin(), vResultText.end(), empty);
	vResultText.erase(itResultText, vResultText.end());
	if (!vResultText.empty())
	{
		if (vResultText[0] == "[yes]\t[no]\t[year]\t[name]\t[path]")
		{
			vResultText.erase(vResultText.begin());
		}
	}
	for (itResultText = vResultText.begin(); itResultText != vResultText.end(); ++itResultText)
	{
		wstring sLine = U8ToW(*itResultText);
		vector<wstring> vLine = Split(sLine, L"\t");
		if (vLine.size() != 5)
		{
			return 1;
		}
		SResult result;
		if (vLine[0] == L"[v]" && trim(vLine[1]).empty())
		{
			result.Exist = true;
		}
		else if (vLine[1] == L"[x]" && trim(vLine[0]).empty())
		{
			result.Exist = false;
		}
		else
		{
			return 1;
		}
		result.Year = SToN32(vLine[2]);
		result.Name = vLine[3];
		result.Path = vLine[4];
		if (!result.Path.empty())
		{
			result.Type = result.Path.substr(0, result.Path.size() - result.Name.size() - 1);
			wstring::size_type uPos = result.Type.find_last_of(L"/\\");
			if (uPos != wstring::npos)
			{
				result.Type.erase(0, uPos + 1);
			}
		}
		m_vResult.push_back(result);
	}
	return 0;
}

void CSwitchGamesXlsx::updateSharedStrings()
{
	for (n32 i = 0; i < static_cast<n32>(m_vSheetName.size()); i++)
	{
		const wstring& sTableName = m_vSheetName[i];
		SSheetInfo& sheetInfo = mSheetInfo[sTableName];
		sheetInfo.Width.clear();
		sheetInfo.Width.resize(sheetInfo.ColumnCount, 9);
		map<n32, map<n32, pair<bool, wstring>>>& mRowColumnText = mTableRowColumnText[sTableName];
		for (n32 nRowIndex = 0; nRowIndex < sheetInfo.RowCount; nRowIndex++)
		{
			map<n32, pair<bool, wstring>>& mColumnText = mRowColumnText[nRowIndex];
			for (n32 nColumnIndex = 0; nColumnIndex < sheetInfo.ColumnCount; nColumnIndex++)
			{
				pair<bool, wstring>& pText = mColumnText[nColumnIndex];
				if (pText.first)
				{
					wstring& sText = pText.second;
					if (static_cast<n32>(sText.size()) > sheetInfo.Width[nColumnIndex])
					{
						sheetInfo.Width[nColumnIndex] = static_cast<n32>(sText.size());
					}
					if (m_bResave)
					{
						if (nRowIndex >= 2)
						{
							m_mSharedStringsIndexNew[sText] = 0;
						}
					}
					else
					{
						m_mSharedStringsIndexNew[sText] = 0;
					}
				}
			}
		}
	}
	for (map<wstring, n32>::const_iterator it = m_mSharedStringsIndexNew.begin(); it != m_mSharedStringsIndexNew.end(); /*it*/)
	{
		if (it->second < 0)
		{
			it = m_mSharedStringsIndexNew.erase(it);
		}
		else
		{
			++it;
		}
	}
	n32 nSharedStringsIndexNew = 0;
	for (map<wstring, n32>::iterator it = m_mSharedStringsIndexNew.begin(); it != m_mSharedStringsIndexNew.end(); ++it)
	{
		it->second = nSharedStringsIndexNew++;
	}
}

bool CSwitchGamesXlsx::matchGreenStyle(const wstring& a_sName) const
{
	for (vector<wregex>::const_iterator itGreen = m_vGreenStyleNamePattern.begin(); itGreen != m_vGreenStyleNamePattern.end(); ++itGreen)
	{
		if (regex_search(a_sName, *itGreen))
		{
			for (vector<wregex>::const_iterator itNotGreen = m_vNotGreenStyleNamePattern.begin(); itNotGreen != m_vNotGreenStyleNamePattern.end(); ++itNotGreen)
			{
				if (regex_search(a_sName, *itNotGreen))
				{
					return false;
				}
			}
			return true;
		}
	}
	return false;
}

int CSwitchGamesXlsx::makePatchTypeFileList()
{
	for (n32 i = 0; i < static_cast<n32>(m_vSheetName.size()); i++)
	{
		const wstring& sType = m_vSheetName[i];
		map<n32, n32>& mRowStyle = m_mTableRowStyle[sType];
		map<n32, map<n32, pair<bool, wstring>>>& mRowColumnText = mTableRowColumnText[sType];
		map<UString, n32> mDirIndex;
		for (map<n32, map<n32, pair<bool, wstring>>>::iterator itRow = mRowColumnText.begin(); itRow != mRowColumnText.end(); ++itRow)
		{
			map<n32, pair<bool, wstring>>& mColumnText = itRow->second;
			if (mColumnText[0].first)
			{
				if (!mDirIndex.insert(make_pair(WToU(mColumnText[0].second), itRow->first)).second)
				{
					return 1;
				}
			}
		}
		n32 nRowIndex = -1;
		UString sRootPath = m_sTableDirName + USTR("/") + WToU(sType);
		UString::size_type uPrefixSize = m_sTableDirName.size() + 1;
		vector<pair<UString, bool>>& vFile = m_vPatchFileList;
		queue<pair<UString, bool>> qDir;
		qDir.push(make_pair(sRootPath, false));
		while (!qDir.empty())
		{
			UString& sParent = qDir.front().first;
			bool bParentStyleIsGreen = qDir.front().second;
#if SDW_PLATFORM == SDW_PLATFORM_WINDOWS
			WIN32_FIND_DATAW ffd;
			HANDLE hFind = INVALID_HANDLE_VALUE;
			wstring sPattern = sParent + L"/*";
			hFind = FindFirstFileW(sPattern.c_str(), &ffd);
			if (hFind != INVALID_HANDLE_VALUE)
			{
				if (sParent == sRootPath)
				{
					m_vPatchTypeList.push_back(sType);
				}
				do
				{
					if ((ffd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
					{
						wstring sFileName = sParent + L"/" + ffd.cFileName;
						vFile.push_back(make_pair(sFileName.substr(uPrefixSize), bParentStyleIsGreen));
					}
					else if ((ffd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0 && wcscmp(ffd.cFileName, L".") != 0 && wcscmp(ffd.cFileName, L"..") != 0)
					{
						wstring sDir = ffd.cFileName;
						map<UString, n32>::const_iterator itDir = mDirIndex.find(sDir);
						if (itDir == mDirIndex.end())
						{
							FindClose(hFind);
							return 1;
						}
						nRowIndex = itDir->second;
						if (nRowIndex < 2)
						{
							FindClose(hFind);
							return 1;
						}
						bool bStyleIsGreen = false;
						if (sParent == sRootPath)
						{
							bStyleIsGreen = mRowStyle[nRowIndex] == kStyleIdGreen;
						}
						sDir = sParent + L"/" + sDir;
						qDir.push(make_pair(sDir, bStyleIsGreen));
					}
				} while (FindNextFileW(hFind, &ffd) != 0);
				FindClose(hFind);
			}
#else
			DIR* pDir = opendir(sParent.c_str());
			if (pDir != nullptr)
			{
				if (sParent == sRootPath)
				{
					m_vPatchTypeList.push_back(WToU(sType));
				}
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
					if (pDirent->d_type == DT_REG)
					{
						string sFileName = sParent + "/" + sName;
						vFile.push_back(make_pair(sFileName.substr(uPrefixSize), bParentStyleIsGreen));
					}
					else if (pDirent->d_type == DT_DIR && strcmp(pDirent->d_name, ".") != 0 && strcmp(pDirent->d_name, "..") != 0)
					{
						string sDir = sName;
						map<UString, n32>::const_iterator itDir = mDirIndex.find(sDir);
						if (itDir == mDirIndex.end())
						{
							closedir(pDir);
							return 1;
						}
						nRowIndex = itDir->second;
						if (nRowIndex < 2)
						{
							closedir(pDir);
							return 1;
						}
						bool bStyleIsGreen = false;
						if (sParent == sRootPath)
						{
							bStyleIsGreen = mRowStyle[nRowIndex] == kStyleIdGreen;
						}
						sDir = sParent + "/" + sDir;
						qDir.push(make_pair(sDir, bStyleIsGreen));
					}
				}
				closedir(pDir);
			}
#endif
			qDir.pop();
		}
	}
	sort(m_vPatchFileList.begin(), m_vPatchFileList.end(), fileListCompare);
	return 0;
}

int CSwitchGamesXlsx::makeRclonePatchBat() const
{
	UString sTableDirName = m_sTableDirName;
#if SDW_PLATFORM == SDW_PLATFORM_WINDOWS
	u32 uMaxPath = sTableDirName.size() + MAX_PATH * 2;
	wchar_t* pTableDirName = new wchar_t[uMaxPath];
	if (_wfullpath(pTableDirName, UToW(sTableDirName).c_str(), uMaxPath) == nullptr)
	{
		return 1;
	}
	sTableDirName = WToU(pTableDirName);
	sTableDirName = Replace(sTableDirName, USTR("\\"), USTR("/"));
#endif
	string sSrcPrefix = UToU8(sTableDirName);
	string sDestPrefix = UToU8(m_sRemoteDirName);
	UString sRcloneCopyBatFileName = m_sTableDirName + USTR("/rclone_0_copy.bat");
	FILE* fp = UFopen(sRcloneCopyBatFileName.c_str(), USTR("wb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	fprintf(fp, "CHCP 65001\r\n");
	fprintf(fp, "PUSHD \"%%~dp0\"\r\n");
	for (vector<UString>::const_iterator it = m_vPatchTypeList.begin(); it != m_vPatchTypeList.end(); ++it)
	{
		string sDirName = UToU8(*it);
		fprintf(fp, "rclone copy \"%s/%s\" \"%s/%s\" --checksum -P --drive-server-side-across-configs\r\n", sSrcPrefix.c_str(), sDirName.c_str(), sDestPrefix.c_str(), sDirName.c_str());
	}
	fprintf(fp, "POPD\r\n");
	fprintf(fp, "PAUSE\r\n");
	fclose(fp);
	UString sRcloneCheckBatFileName = m_sTableDirName + USTR("/rclone_1_check.bat");
	fp = UFopen(sRcloneCheckBatFileName.c_str(), USTR("wb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	fprintf(fp, "CHCP 65001\r\n");
	fprintf(fp, "PUSHD \"%%~dp0\"\r\n");
	for (vector<UString>::const_iterator it = m_vPatchTypeList.begin(); it != m_vPatchTypeList.end(); ++it)
	{
		string sDirName = UToU8(*it);
		fprintf(fp, "rclone check \"%s/%s\" \"%s/%s\" --one-way --checksum -P --drive-server-side-across-configs || PAUSE\r\n", sSrcPrefix.c_str(), sDirName.c_str(), sDestPrefix.c_str(), sDirName.c_str());
	}
	fprintf(fp, "POPD\r\n");
	fprintf(fp, "PAUSE\r\n");
	fclose(fp);
	return 0;
}

int CSwitchGamesXlsx::makeBaiduPCSGoPatchBat() const
{
	UString sTableDirName = m_sTableDirName;
#if SDW_PLATFORM == SDW_PLATFORM_WINDOWS
	u32 uMaxPath = sTableDirName.size() + MAX_PATH * 2;
	wchar_t* pTableDirName = new wchar_t[uMaxPath];
	if (_wfullpath(pTableDirName, UToW(sTableDirName).c_str(), uMaxPath) == nullptr)
	{
		return 1;
	}
	sTableDirName = WToU(pTableDirName);
	sTableDirName = Replace(sTableDirName, USTR("\\"), USTR("/"));
#endif
	string sSrcPrefix = UToU8(sTableDirName);
	string sDestPrefix = UToU8(m_sRemoteDirName);
	UString sBaiduPCSGoUploadTxtFileName;
	UString sBaiduPCSGoUploadBatFileName;
	if (m_bStyleIsGreen)
	{
		sBaiduPCSGoUploadTxtFileName = m_sTableDirName + USTR("/baidupcs-go_0_0_upload_green.txt");
		sBaiduPCSGoUploadBatFileName = m_sTableDirName + USTR("/baidupcs-go_0_1_upload_green.bat");
	}
	else
	{
		sBaiduPCSGoUploadTxtFileName = m_sTableDirName + USTR("/baidupcs-go_2_0_upload.txt");
		sBaiduPCSGoUploadBatFileName = m_sTableDirName + USTR("/baidupcs-go_2_1_upload.bat");
	}
	FILE* fp = UFopen(sBaiduPCSGoUploadTxtFileName.c_str(), USTR("wb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	fprintf(fp, "su %s\r\n", UToU8(m_sBaiduUserId).c_str());
	for (vector<pair<UString, bool>>::const_iterator it = m_vPatchFileList.begin(); it != m_vPatchFileList.end(); ++it)
	{
		if (it->second == m_bStyleIsGreen)
		{
			string sFileName = UToU8(it->first);
			string sDirName = ".";
			string::size_type uPos = sFileName.rfind("/");
			if (uPos != string::npos)
			{
				sDirName = sFileName.substr(0, uPos);
			}
			fprintf(fp, "rm \"%s/%s\"\r\n", sDestPrefix.c_str(), sFileName.c_str());
			fprintf(fp, "upload \"%s/%s\" \"%s/%s\"\r\n", sSrcPrefix.c_str(), sFileName.c_str(), sDestPrefix.c_str(), sDirName.c_str());
			fprintf(fp, "fixmd5 \"%s/%s\"\r\n", sDestPrefix.c_str(), sFileName.c_str());
		}
	}
	fclose(fp);
	fp = UFopen(sBaiduPCSGoUploadBatFileName.c_str(), USTR("wb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	fprintf(fp, "CHCP 65001\r\n");
	fprintf(fp, "PUSHD \"%%~dp0\"\r\n");
	fprintf(fp, "BaiduPCS-Go < \"%s\"\r\n", UToU8(sBaiduPCSGoUploadTxtFileName.substr(m_sTableDirName.size() + 1)).c_str());
	fprintf(fp, "POPD\r\n");
	fprintf(fp, "PAUSE\r\n");
	fclose(fp);
	string sBaiduPCSGoMd5TxtFileName;
	UString sBaiduPCSGoSumMetaTxtFileName;
	UString sBaiduPCSGoSumMetaCheckBatFileName;
	if (m_bStyleIsGreen)
	{
		sBaiduPCSGoMd5TxtFileName = "baidupcs-go_1_0_md5_green.txt";
		sBaiduPCSGoSumMetaTxtFileName = m_sTableDirName + USTR("/baidupcs-go_1_0_sum_meta_green.txt");
		sBaiduPCSGoSumMetaCheckBatFileName = m_sTableDirName + USTR("/baidupcs-go_1_1_sum_meta_check_green.bat");
	}
	else
	{
		sBaiduPCSGoMd5TxtFileName = "baidupcs-go_3_0_md5.txt";
		sBaiduPCSGoSumMetaTxtFileName = m_sTableDirName + USTR("/baidupcs-go_3_0_sum_meta.txt");
		sBaiduPCSGoSumMetaCheckBatFileName = m_sTableDirName + USTR("/baidupcs-go_3_1_sum_meta_check.bat");
	}
	fp = UFopen(sBaiduPCSGoSumMetaTxtFileName.c_str(), USTR("wb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	fprintf(fp, "su %s\r\n", UToU8(m_sBaiduUserId).c_str());
	fprintf(fp, "who\r\n");
	for (vector<pair<UString, bool>>::const_iterator it = m_vPatchFileList.begin(); it != m_vPatchFileList.end(); ++it)
	{
		if (it->second == m_bStyleIsGreen)
		{
			string sFileName = UToU8(it->first);
			fprintf(fp, "sumfile \"%s/%s\"\r\n", sSrcPrefix.c_str(), sFileName.c_str());
			fprintf(fp, "meta \"%s/%s\"\r\n", sDestPrefix.c_str(), sFileName.c_str());
		}
	}
	fclose(fp);
	fp = UFopen(sBaiduPCSGoSumMetaCheckBatFileName.c_str(), USTR("wb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	fprintf(fp, "CHCP 65001\r\n");
	fprintf(fp, "PUSHD \"%%~dp0\"\r\n");
	fprintf(fp, "BaiduPCS-Go < \"%s\" > \"%s\"\r\n", UToU8(sBaiduPCSGoSumMetaTxtFileName.substr(m_sTableDirName.size() + 1)).c_str(), sBaiduPCSGoMd5TxtFileName.c_str());
	fprintf(fp, "CheckBaiduPCSGoMd5 \"%s\" || PAUSE\r\n", sBaiduPCSGoMd5TxtFileName.c_str());
	fprintf(fp, "POPD\r\n");
	fprintf(fp, "PAUSE\r\n");
	fclose(fp);
	return 0;
}
