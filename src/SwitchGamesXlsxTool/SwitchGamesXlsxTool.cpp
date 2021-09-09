#include <sdw.h>
#include <tinyxml2.h>

enum EFillId
{
	kFillIdNone = 0,
	kFillIdGray125 = 1,
	kFillIdRed = 2,
	kFillIdBlue = 3,
	kFillIdGreen = 4,
	kFillIdYellow = 5,
	kFillIdGold = 6
};

enum EStyleId
{
	kStyleIdNoneNoBorder = 0,
	kStyleIdNone = 1,
	kStyleIdRed = 2,
	kStyleIdBlueBold = 3,
	kStyleIdBlue = 4,
	kStyleIdGreen = 5,
	kStyleIdYellow = 6,
	kStyleIdGold = 7
};

struct SSheetInfo
{
	n32 RowCount;
	n32 ColumnCount;
	bool TabSelected;
	n32 TopLeftCellRowIndex;
	n32 TopLeftCellColumnIndex;
	n32 ActiveCellRowIndex;
	n32 ActiveCellColumnIndex;
	vector<n32> Width;
	SSheetInfo()
		: RowCount(0)
		, ColumnCount(0)
		, TabSelected(false)
		, TopLeftCellRowIndex(0)
		, TopLeftCellColumnIndex(0)
		, ActiveCellRowIndex(0)
		, ActiveCellColumnIndex(0)
	{
	}
};

string validateText(const string& a_sText)
{
	string sText = Replace(a_sText, '&', "&amp;");
	sText = Replace(sText, '<', "&lt;");
	sText = Replace(sText, '>', "&gt;");
	return sText;
}

string validateValue(const string& a_sValue)
{
	string sValue = Replace(a_sValue, '\"', "&quot;");
	return sValue;
}

bool travelElement(const tinyxml2::XMLElement* a_pRootElement, string& a_sXml)
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

string encodePosition(n32 a_nRowIndex, n32 a_nColumnIndex)
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

bool copyFile(const UString& a_sDestFileName, const UString& a_sSrcFileName)
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

int resaveXlsx(const UString& a_sXlsxDirName)
{
	bool bCompact = false;
	UString sModuleDirName = UGetModuleDirName();

	vector<wstring> vSheetName;
	set<wstring> sSheetName;

#pragma region config
	UString sConfigPath = sModuleDirName + USTR("/Switch Games.xml");
	FILE* fp = UFopen(sConfigPath.c_str(), USTR("rb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	tinyxml2::XMLDocument configDoc;
	tinyxml2::XMLError eError = configDoc.LoadFile(fp);
	fclose(fp);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		return 1;
	}
	tinyxml2::XMLElement* pConfigDocConfig = configDoc.FirstChildElement("config");
	tinyxml2::XMLElement* pConfigDocConfigSheets = pConfigDocConfig->FirstChildElement("sheets");
	for (tinyxml2::XMLElement* pSheetsSheet = pConfigDocConfigSheets->FirstChildElement("sheet"); pSheetsSheet != nullptr; pSheetsSheet = pSheetsSheet->NextSiblingElement("sheet"))
	{
		wstring sSheetText;
		const char* pSheetText = pSheetsSheet->GetText();
		if (pSheetText != nullptr)
		{
			sSheetText = U8ToW(pSheetText);
		}
		vSheetName.push_back(sSheetText);
		sSheetName.insert(sSheetText);
	}
#pragma endregion config

	n32 nSheetCountOld = 0;
	map<wstring, n32> mSheetIndexOld;
	string sCreator;
	string sLastModifiedBy;
	string sLastModifiedDateTime;
	n32 nActiveTabOld = 0;
	n32 nActiveTabNew = 0;
	n32 nSstCount = -1;
	map<n32, wstring> mSharedStrings;
	map<wstring, n32> mSharedStringsIndexNew;
	map<n32, n32> mStyleOldToNew;
	map<wstring, SSheetInfo> mSheetInfo;
	map<wstring, map<n32, n32>> mTableRowStyle;
	map<wstring, map<n32, map<n32, pair<bool, wstring>>>> mTableRowColumnText;

#pragma region read /xl/workbook.xml
	// read /xl/workbook.xml
	UString sWorkbookXmlPath = a_sXlsxDirName + USTR("/xl/workbook.xml");
	fp = UFopen(sWorkbookXmlPath.c_str(), USTR("rb"), false);
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
	tinyxml2::XMLDocument workbookXmlDoc;
	eError = workbookXmlDoc.LoadFile(fp);
	fclose(fp);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		return 1;
	}
	if (!StartWith(sWorkbookXml, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"))
	{
		return 1;
	}
	tinyxml2::XMLElement* pWorkbookXmlDocWorkbook = workbookXmlDoc.RootElement();
	if (pWorkbookXmlDocWorkbook == nullptr)
	{
		return 1;
	}
	string sWorkbookXmlDocWorkbookName = pWorkbookXmlDocWorkbook->Name();
	if (sWorkbookXmlDocWorkbookName != "workbook")
	{
		return 1;
	}
	tinyxml2::XMLElement* pWorkbookXmlDocWorkbookBookViews = pWorkbookXmlDocWorkbook->FirstChildElement("bookViews");
	if (pWorkbookXmlDocWorkbookBookViews != nullptr)
	{
		tinyxml2::XMLElement* pBookViewsWorkbookView = pWorkbookXmlDocWorkbookBookViews->FirstChildElement("workbookView");
		if (pBookViewsWorkbookView != nullptr)
		{
			const char* pWorkbookViewActiveTab = pBookViewsWorkbookView->Attribute("activeTab");
			if (pWorkbookViewActiveTab != nullptr)
			{
				nActiveTabOld = SToN32(pWorkbookViewActiveTab);
			}
		}
	}
#pragma endregion read /xl/workbook.xml

#pragma region read /xl/sharedStrings.xml
	// read /xl/sharedStrings.xml
	UString sSharedStringsXmlPath = a_sXlsxDirName + USTR("/xl/sharedStrings.xml");
	fp = UFopen(sSharedStringsXmlPath.c_str(), USTR("rb"), false);
	if (fp == nullptr)
	{
		// do nothing
	}
	else
	{
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
		tinyxml2::XMLDocument sharedStringsXmlDoc;
		eError = sharedStringsXmlDoc.LoadFile(fp);
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
		const tinyxml2::XMLElement* pSharedStringsDocSst = sharedStringsXmlDoc.RootElement();
		if (pSharedStringsDocSst == nullptr)
		{
			return 1;
		}
		string sSstName = pSharedStringsDocSst->Name();
		if (sSstName != "sst")
		{
			return 1;
		}
		n32 nSstUniqueCount = -1;
		sSharedStringsXml += "<";
		sSharedStringsXml += pSharedStringsDocSst->Name();
		for (const tinyxml2::XMLAttribute* pSstAttribute = pSharedStringsDocSst->FirstAttribute(); pSstAttribute != nullptr; pSstAttribute = pSstAttribute->Next())
		{
			string sAttributeName = pSstAttribute->Name();
			string sAttributeValue = pSstAttribute->Value();
			if (sAttributeName == "count")
			{
				nSstCount = SToN32(sAttributeValue);
			}
			else if (sAttributeName == "uniqueCount")
			{
				nSstUniqueCount = SToN32(sAttributeValue);
			}
			sAttributeValue = validateValue(sAttributeValue);
			sSharedStringsXml += " " + sAttributeName + "=\"" + sAttributeValue + "\"";
		}
		sSharedStringsXml += ">";
		if (nSstCount < 0 || nSstUniqueCount < 0)
		{
			return 1;
		}
		n32 nSstSiCount = 0;
		for (const tinyxml2::XMLNode* pSstChild = pSharedStringsDocSst->FirstChild(); pSstChild != nullptr; pSstChild = pSstChild->NextSibling())
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
				if (!mSharedStrings.insert(make_pair(nSstSiCount, U8ToW(sText))).second)
				{
					return 1;
				}
				nSstSiCount++;
			}
		}
		sSharedStringsXml += "</";
		sSharedStringsXml += pSharedStringsDocSst->Name();
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
	}
	for (map<n32, wstring>::iterator it = mSharedStrings.begin(); it != mSharedStrings.end(); ++it)
	{
		const wstring& sText = it->second;
		mSharedStringsIndexNew.insert(make_pair(sText, -1));
	}
	mSharedStringsIndexNew[L"Release Name"] = 0;
	mSharedStringsIndexNew[L"File Name"] = 0;
	mSharedStringsIndexNew[L"Comment"] = 0;
	mSharedStringsIndexNew[L"Comment2"] = 0;
	mSharedStringsIndexNew[L".."] = 0;
	mSharedStringsIndexNew[L"nfo LF/sfv CRLF"] = 0;
#pragma endregion read /xl/sharedStrings.xml

#pragma region read /xl/styles.xml
	// read /xl/styles.xml
	UString sStylesXmlPath = a_sXlsxDirName + USTR("/xl/styles.xml");
	fp = UFopen(sStylesXmlPath.c_str(), USTR("rb"), false);
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
	tinyxml2::XMLDocument stylesXmlDoc;
	eError = stylesXmlDoc.LoadFile(fp);
	fclose(fp);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		return 1;
	}
	if (!StartWith(sStylesXml, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"))
	{
		return 1;
	}
	tinyxml2::XMLElement* pStylesXmlDocStyleSheet = stylesXmlDoc.RootElement();
	if (pStylesXmlDocStyleSheet == nullptr)
	{
		return 1;
	}
	string sStylesXmlDocStyleSheetName = pStylesXmlDocStyleSheet->Name();
	if (sStylesXmlDocStyleSheetName != "styleSheet")
	{
		return 1;
	}
	tinyxml2::XMLElement* pStylesXmlDocStyleSheetFills = pStylesXmlDocStyleSheet->FirstChildElement("fills");
	if (pStylesXmlDocStyleSheetFills == nullptr)
	{
		return 1;
	}
	map<n32, n32> mFillOldToNew;
	n32 nFillIndexOld = 0;
	for (tinyxml2::XMLElement* pFillsFill = pStylesXmlDocStyleSheetFills->FirstChildElement("fill"); pFillsFill != nullptr; pFillsFill = pFillsFill->NextSiblingElement("fill"))
	{
		tinyxml2::XMLElement* pFillPatternFill = pFillsFill->FirstChildElement("patternFill");
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
			tinyxml2::XMLElement* pPatternFillFgColor = pFillPatternFill->FirstChildElement("fgColor");
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
	tinyxml2::XMLElement* pStylesXmlDocStyleSheetCellXfs = pStylesXmlDocStyleSheet->FirstChildElement("cellXfs");
	if (pStylesXmlDocStyleSheetCellXfs == nullptr)
	{
		return 1;
	}
	n32 nXfIndexOld = 0;
	for (tinyxml2::XMLElement* pCellXfsXf = pStylesXmlDocStyleSheetCellXfs->FirstChildElement("xf"); pCellXfsXf != nullptr; pCellXfsXf = pCellXfsXf->NextSiblingElement("xf"))
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
			mStyleOldToNew[nXfIndexOld++] = kStyleIdNone;
			break;
		case kFillIdRed:
			mStyleOldToNew[nXfIndexOld++] = kStyleIdRed;
			break;
		case kFillIdBlue:
			mStyleOldToNew[nXfIndexOld++] = kStyleIdBlue;
			break;
		case kFillIdGreen:
			mStyleOldToNew[nXfIndexOld++] = kStyleIdGreen;
			break;
		case kFillIdYellow:
			mStyleOldToNew[nXfIndexOld++] = kStyleIdYellow;
			break;
		case kFillIdGold:
			mStyleOldToNew[nXfIndexOld++] = kStyleIdGold;
			break;
		default:
			mStyleOldToNew[nXfIndexOld++] = kStyleIdRed;
			break;
		}
	}
#pragma endregion read /xl/styles.xml

#pragma region /_rels/.rels
	// /_rels/.rels
	UString s_relsPath = a_sXlsxDirName + USTR("/_rels/.rels");
	UString sSwitchGames_relsPath = sModuleDirName + USTR("/Switch Games/_rels/.rels");
	if (!copyFile(s_relsPath, sSwitchGames_relsPath))
	{
		return 1;
	}
#pragma endregion /_rels/.rels

#pragma region /docProps/app.xml
	// /docProps/app.xml
	UString sAppXmlPath = a_sXlsxDirName + USTR("/docProps/app.xml");
	fp = UFopen(sAppXmlPath.c_str(), USTR("rb"), false);
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
	tinyxml2::XMLDocument appXmlDoc;
	eError = appXmlDoc.LoadFile(fp);
	fclose(fp);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		return 1;
	}
	if (!StartWith(sAppXml, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"))
	{
		return 1;
	}
	tinyxml2::XMLElement* pAppXmlDocProperties = appXmlDoc.RootElement();
	if (pAppXmlDocProperties == nullptr)
	{
		return 1;
	}
	string sAppXmlDocPropertiesName = pAppXmlDocProperties->Name();
	if (sAppXmlDocPropertiesName != "Properties")
	{
		return 1;
	}
	tinyxml2::XMLElement* pAppXmlDocPropertiesHeadingPairs = pAppXmlDocProperties->FirstChildElement("HeadingPairs");
	if (pAppXmlDocPropertiesHeadingPairs == nullptr)
	{
		return 1;
	}
	tinyxml2::XMLElement* pAppXmlDocPropertiesHeadingPairsVtVector = pAppXmlDocPropertiesHeadingPairs->FirstChildElement("vt:vector");
	if (pAppXmlDocPropertiesHeadingPairsVtVector == nullptr)
	{
		return 1;
	}
	tinyxml2::XMLElement* pAppXmlDocPropertiesHeadingPairsVtVectorVtVariant = pAppXmlDocPropertiesHeadingPairsVtVector->FirstChildElement("vt:variant");
	if (pAppXmlDocPropertiesHeadingPairsVtVectorVtVariant == nullptr)
	{
		return 1;
	}
	pAppXmlDocPropertiesHeadingPairsVtVectorVtVariant = pAppXmlDocPropertiesHeadingPairsVtVectorVtVariant->NextSiblingElement("vt:variant");
	if (pAppXmlDocPropertiesHeadingPairsVtVectorVtVariant == nullptr)
	{
		return 1;
	}
	tinyxml2::XMLElement* pAppXmlDocPropertiesHeadingPairsVtVectorVtVariantVtI4 = pAppXmlDocPropertiesHeadingPairsVtVectorVtVariant->FirstChildElement("vt:i4");
	if (pAppXmlDocPropertiesHeadingPairsVtVectorVtVariantVtI4 == nullptr)
	{
		return 1;
	}
	nSheetCountOld = pAppXmlDocPropertiesHeadingPairsVtVectorVtVariantVtI4->IntText();
	tinyxml2::XMLElement* pAppXmlDocPropertiesTitlesOfParts = pAppXmlDocProperties->FirstChildElement("TitlesOfParts");
	if (pAppXmlDocPropertiesTitlesOfParts == nullptr)
	{
		return 1;
	}
	tinyxml2::XMLElement* pAppXmlDocPropertiesTitlesOfPartsVtVector = pAppXmlDocPropertiesTitlesOfParts->FirstChildElement("vt:vector");
	if (pAppXmlDocPropertiesTitlesOfPartsVtVector == nullptr)
	{
		return 1;
	}
	if (pAppXmlDocPropertiesTitlesOfPartsVtVector->IntAttribute("size") != nSheetCountOld)
	{
		return 1;
	}
	n32 nSheetIndexOld = 1;
	for (tinyxml2::XMLElement* pVtVectorVtLpstr = pAppXmlDocPropertiesTitlesOfPartsVtVector->FirstChildElement("vt:lpstr"); pVtVectorVtLpstr != nullptr; pVtVectorVtLpstr = pVtVectorVtLpstr->NextSiblingElement("vt:lpstr"))
	{
		wstring sVtLpstrText;
		const char* pVtLpstrText = pVtVectorVtLpstr->GetText();
		if (pVtLpstrText != nullptr)
		{
			sVtLpstrText = U8ToW(pVtLpstrText);
		}
		if (sSheetName.find(sVtLpstrText) == sSheetName.end())
		{
			continue;
		}
		if (nSheetIndexOld == nActiveTabOld + 1)
		{
			for (n32 i = 0; i < static_cast<n32>(vSheetName.size()); i++)
			{
				if (vSheetName[i] == sVtLpstrText)
				{
					nActiveTabNew = i;
				}
			}
		}
		mSheetIndexOld[sVtLpstrText] = nSheetIndexOld++;
	}
	UString sSwitchGamesAppXmlPath = sModuleDirName + USTR("/Switch Games/docProps/app.xml");
	fp = UFopen(sSwitchGamesAppXmlPath.c_str(), USTR("rb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	eError = appXmlDoc.LoadFile(fp);
	fclose(fp);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		return 1;
	}
	pAppXmlDocProperties = appXmlDoc.RootElement();
	pAppXmlDocPropertiesHeadingPairs = pAppXmlDocProperties->FirstChildElement("HeadingPairs");
	pAppXmlDocPropertiesHeadingPairsVtVector = pAppXmlDocPropertiesHeadingPairs->FirstChildElement("vt:vector");
	pAppXmlDocPropertiesHeadingPairsVtVectorVtVariant = pAppXmlDocPropertiesHeadingPairsVtVector->FirstChildElement("vt:variant");
	pAppXmlDocPropertiesHeadingPairsVtVectorVtVariant = pAppXmlDocPropertiesHeadingPairsVtVectorVtVariant->NextSiblingElement("vt:variant");
	pAppXmlDocPropertiesHeadingPairsVtVectorVtVariantVtI4 = pAppXmlDocPropertiesHeadingPairsVtVectorVtVariant->FirstChildElement("vt:i4");
	pAppXmlDocPropertiesHeadingPairsVtVectorVtVariantVtI4->SetText(static_cast<n32>(vSheetName.size()));
	pAppXmlDocPropertiesTitlesOfParts = pAppXmlDocProperties->FirstChildElement("TitlesOfParts");
	pAppXmlDocPropertiesTitlesOfPartsVtVector = pAppXmlDocPropertiesTitlesOfParts->FirstChildElement("vt:vector");
	pAppXmlDocPropertiesTitlesOfPartsVtVector->SetAttribute("size", static_cast<n32>(vSheetName.size()));
	pAppXmlDocPropertiesTitlesOfPartsVtVector->DeleteChildren();
	for (n32 i = 0; i < static_cast<n32>(vSheetName.size()); i++)
	{
		tinyxml2::XMLElement* pVtVectorVtLpstr = pAppXmlDocPropertiesTitlesOfPartsVtVector->InsertNewChildElement("vt:lpstr");
		pVtVectorVtLpstr->SetText(WToU8(vSheetName[i]).c_str());
	}
	fp = UFopen(sAppXmlPath.c_str(), USTR("wb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	eError = appXmlDoc.SaveFile(fp, bCompact);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		fclose(fp);
		return 1;
	}
	fclose(fp);
#pragma endregion /docProps/app.xml

#pragma region /docProps/core.xml
	// /docProps/core.xml
	UString sCoreXmlPath = a_sXlsxDirName + USTR("/docProps/core.xml");
	fp = UFopen(sCoreXmlPath.c_str(), USTR("rb"), false);
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
	tinyxml2::XMLDocument coreXmlDoc;
	eError = coreXmlDoc.LoadFile(fp);
	fclose(fp);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		return 1;
	}
	if (!StartWith(sCoreXml, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"))
	{
		return 1;
	}
	tinyxml2::XMLElement* pCoreXmlDocCpCoreProperties = coreXmlDoc.RootElement();
	if (pCoreXmlDocCpCoreProperties == nullptr)
	{
		return 1;
	}
	string sCoreXmlDocCpCorePropertiesName = pCoreXmlDocCpCoreProperties->Name();
	if (sCoreXmlDocCpCorePropertiesName != "cp:coreProperties")
	{
		return 1;
	}
	tinyxml2::XMLElement* pCoreXmlDocCpCorePropertiesDcCreator = pCoreXmlDocCpCoreProperties->FirstChildElement("dc:creator");
	if (pCoreXmlDocCpCorePropertiesDcCreator == nullptr)
	{
		return 1;
	}
	const char* pCoreXmlDocCpCorePropertiesDcCreatorText = pCoreXmlDocCpCorePropertiesDcCreator->GetText();
	if (pCoreXmlDocCpCorePropertiesDcCreatorText != nullptr)
	{
		sCreator = pCoreXmlDocCpCorePropertiesDcCreatorText;
	}
	tinyxml2::XMLElement* pCoreXmlDocCpCorePropertiesCpLastModifiedBy = pCoreXmlDocCpCoreProperties->FirstChildElement("cp:lastModifiedBy");
	if (pCoreXmlDocCpCorePropertiesCpLastModifiedBy == nullptr)
	{
		return 1;
	}
	const char* pCoreXmlDocCpCorePropertiesCpLastModifiedByText = pCoreXmlDocCpCorePropertiesCpLastModifiedBy->GetText();
	if (pCoreXmlDocCpCorePropertiesCpLastModifiedByText != nullptr)
	{
		sLastModifiedBy = pCoreXmlDocCpCorePropertiesCpLastModifiedByText;
	}
	tinyxml2::XMLElement* pCoreXmlDocCpCorePropertiesDctermsModified = pCoreXmlDocCpCoreProperties->FirstChildElement("dcterms:modified");
	if (pCoreXmlDocCpCorePropertiesDctermsModified == nullptr)
	{
		return 1;
	}
	const char* pCoreXmlDocCpCorePropertiesDctermsModifiedText = pCoreXmlDocCpCorePropertiesDctermsModified->GetText();
	if (pCoreXmlDocCpCorePropertiesDctermsModifiedText != nullptr)
	{
		sLastModifiedDateTime = pCoreXmlDocCpCorePropertiesDctermsModifiedText;
	}
	UString sSwitchGamesCoreXmlPath = sModuleDirName + USTR("/Switch Games/docProps/core.xml");
	fp = UFopen(sSwitchGamesCoreXmlPath.c_str(), USTR("rb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	eError = coreXmlDoc.LoadFile(fp);
	fclose(fp);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		return 1;
	}
	pCoreXmlDocCpCoreProperties = coreXmlDoc.RootElement();
	pCoreXmlDocCpCorePropertiesDcCreator = pCoreXmlDocCpCoreProperties->FirstChildElement("dc:creator");
	pCoreXmlDocCpCorePropertiesDcCreator->SetText(sCreator.c_str());
	pCoreXmlDocCpCorePropertiesCpLastModifiedBy = pCoreXmlDocCpCoreProperties->FirstChildElement("cp:lastModifiedBy");
	pCoreXmlDocCpCorePropertiesCpLastModifiedBy->SetText(sLastModifiedBy.c_str());
	pCoreXmlDocCpCorePropertiesDctermsModified = pCoreXmlDocCpCoreProperties->FirstChildElement("dcterms:modified");
	pCoreXmlDocCpCorePropertiesDctermsModified->SetText(sLastModifiedDateTime.c_str());
	fp = UFopen(sCoreXmlPath.c_str(), USTR("wb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	eError = coreXmlDoc.SaveFile(fp, bCompact);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		fclose(fp);
		return 1;
	}
	fclose(fp);
#pragma endregion /docProps/core.xml

#pragma region /xl/_rels/workbook.xml.rels
	// /xl/_rels/workbook.xml.rels
	UString sWorkbookXmlRelsPath = a_sXlsxDirName + USTR("/xl/_rels/workbook.xml.rels");
	tinyxml2::XMLDocument workbookXmlRelsDoc;
	UString sSwitchGamesWorkbookXmlRelsPath = sModuleDirName + USTR("/Switch Games/xl/_rels/workbook.xml.rels");
	fp = UFopen(sSwitchGamesWorkbookXmlRelsPath.c_str(), USTR("rb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	eError = workbookXmlRelsDoc.LoadFile(fp);
	fclose(fp);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		return 1;
	}
	tinyxml2::XMLElement* pWorkbookXmlRelsDocRelationships = workbookXmlRelsDoc.RootElement();
	pWorkbookXmlRelsDocRelationships->DeleteChildren();
	tinyxml2::XMLElement* pWorkbookXmlRelsDocRelationshipsRelationship = pWorkbookXmlRelsDocRelationships->InsertNewChildElement("Relationship");
	pWorkbookXmlRelsDocRelationshipsRelationship->SetAttribute("Id", Format("rId%d", static_cast<n32>(vSheetName.size()) + 3).c_str());
	pWorkbookXmlRelsDocRelationshipsRelationship->SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
	pWorkbookXmlRelsDocRelationshipsRelationship->SetAttribute("Target", "sharedStrings.xml");
	pWorkbookXmlRelsDocRelationshipsRelationship = pWorkbookXmlRelsDocRelationships->InsertNewChildElement("Relationship");
	pWorkbookXmlRelsDocRelationshipsRelationship->SetAttribute("Id", Format("rId%d", static_cast<n32>(vSheetName.size()) + 2).c_str());
	pWorkbookXmlRelsDocRelationshipsRelationship->SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
	pWorkbookXmlRelsDocRelationshipsRelationship->SetAttribute("Target", "styles.xml");
	pWorkbookXmlRelsDocRelationshipsRelationship = pWorkbookXmlRelsDocRelationships->InsertNewChildElement("Relationship");
	pWorkbookXmlRelsDocRelationshipsRelationship->SetAttribute("Id", Format("rId%d", static_cast<n32>(vSheetName.size()) + 1).c_str());
	pWorkbookXmlRelsDocRelationshipsRelationship->SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme");
	pWorkbookXmlRelsDocRelationshipsRelationship->SetAttribute("Target", "theme/theme1.xml");
	for (n32 i = static_cast<n32>(vSheetName.size()); i > 0; i--)
	{
		pWorkbookXmlRelsDocRelationshipsRelationship = pWorkbookXmlRelsDocRelationships->InsertNewChildElement("Relationship");
		pWorkbookXmlRelsDocRelationshipsRelationship->SetAttribute("Id", Format("rId%d", i).c_str());
		pWorkbookXmlRelsDocRelationshipsRelationship->SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
		pWorkbookXmlRelsDocRelationshipsRelationship->SetAttribute("Target", Format("worksheets/sheet%d.xml", i).c_str());
	}
	fp = UFopen(sWorkbookXmlRelsPath.c_str(), USTR("wb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	eError = workbookXmlRelsDoc.SaveFile(fp, bCompact);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		fclose(fp);
		return 1;
	}
	fclose(fp);
#pragma endregion /xl/_rels/workbook.xml.rels

#pragma region /xl/theme/theme1.xml
	// /xl/theme/theme1.xml
	UString sTheme1XmlPath = a_sXlsxDirName + USTR("/xl/theme/theme1.xml");
	UString sSwitchGamesTheme1XmlPath = sModuleDirName + USTR("/Switch Games/xl/theme/theme1.xml");
	if (!copyFile(sTheme1XmlPath, sSwitchGamesTheme1XmlPath))
	{
		return 1;
	}
#pragma endregion /xl/theme/theme1.xml

#pragma region /xl/worksheets/sheet%d.xml
	// /xl/worksheets/sheet%d.xml
	n32 nSheetStringCount = 0;
	for (n32 i = 0; i < static_cast<n32>(vSheetName.size()); i++)
	{
		const wstring& sTableName = vSheetName[i];
		SSheetInfo& sheetInfo = mSheetInfo[sTableName];
		sheetInfo.TabSelected = i == nActiveTabNew;
		map<n32, n32>& mRowStyle = mTableRowStyle[sTableName];
		n32 nSheetIndex = mSheetIndexOld[sTableName];
		if (nSheetIndex < 1)
		{
			sheetInfo.RowCount = 2;
			sheetInfo.ColumnCount = 4;
			sheetInfo.Width.resize(4, 9);
			continue;
		}
		UString sSheetXmlPath = a_sXlsxDirName + Format(USTR("/xl/worksheets/sheet%d.xml"), nSheetIndex);
		fp = UFopen(sSheetXmlPath.c_str(), USTR("rb"), false);
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
		tinyxml2::XMLDocument sheetXmlDoc;
		eError = sheetXmlDoc.LoadFile(fp);
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
		const tinyxml2::XMLElement* pSheetXmlDocWorksheet = sheetXmlDoc.RootElement();
		if (pSheetXmlDocWorksheet == nullptr)
		{
			return 1;
		}
		string sWorksheetName = pSheetXmlDocWorksheet->Name();
		if (sWorksheetName != "worksheet")
		{
			return 1;
		}
		sSheetXml += "<";
		sSheetXml += pSheetXmlDocWorksheet->Name();
		for (const tinyxml2::XMLAttribute* pWorksheetAttribute = pSheetXmlDocWorksheet->FirstAttribute(); pWorksheetAttribute != nullptr; pWorksheetAttribute = pWorksheetAttribute->Next())
		{
			string sAttributeName = pWorksheetAttribute->Name();
			string sAttributeValue = validateValue(pWorksheetAttribute->Value());
			sSheetXml += " " + sAttributeName + "=\"" + sAttributeValue + "\"";
		}
		sSheetXml += ">";
		for (const tinyxml2::XMLNode* pWorksheetChild = pSheetXmlDocWorksheet->FirstChild(); pWorksheetChild != nullptr; pWorksheetChild = pWorksheetChild->NextSibling())
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
								nRowStyleNew = mStyleOldToNew[nRowStyleOld];
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
											nRowStyleNew = mStyleOldToNew[nRowStyleOld];
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
													map<n32, wstring>::const_iterator itSharedStrings = mSharedStrings.find(nStringIndex);
													if (itSharedStrings == mSharedStrings.end())
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
											if (nRowIndex >= 2)
											{
												mSharedStringsIndexNew[sStmtW] = 0;
											}
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
		sSheetXml += pSheetXmlDocWorksheet->Name();
		sSheetXml += ">";
		if (sSheetXml != sSheetXml0)
		{
			return 1;
		}
	}
	if (nSheetStringCount > nSstCount)
	{
		return 1;
	}
	for (vector<wstring>::iterator it = vSheetName.begin(); it != vSheetName.end(); ++it)
	{
		wstring& sTableName = *it;
		SSheetInfo& sheetInfo = mSheetInfo[sTableName];
		map<n32, n32>& mRowStyle = mTableRowStyle[sTableName];
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
		mRowColumnText[1][3] = make_pair(false, L"");
		sheetInfo.RowCount = max<n32>(2, sheetInfo.RowCount);
		sheetInfo.ColumnCount = max<n32>(4, sheetInfo.ColumnCount);
		for (n32 nRowIndex = 0; nRowIndex < 2; nRowIndex++)
		{
			for (n32 nColumnIndex = 0; nColumnIndex < 4; nColumnIndex++)
			{
				wstring& sStmtW = mRowColumnText[nRowIndex][nColumnIndex].second;
				if (static_cast<n32>(sStmtW.size()) > sheetInfo.Width[nColumnIndex])
				{
					sheetInfo.Width[nColumnIndex] = static_cast<n32>(sStmtW.size());
				}
			}
		}
		mRowColumnText[1].erase(3);
	}
	for (map<wstring, n32>::iterator it = mSharedStringsIndexNew.begin(); it != mSharedStringsIndexNew.end(); /*it*/)
	{
		if (it->second < 0)
		{
			it = mSharedStringsIndexNew.erase(it);
		}
		else
		{
			++it;
		}
	}
	n32 nSharedStringsIndexNew = 0;
	for (map<wstring, n32>::iterator it = mSharedStringsIndexNew.begin(); it != mSharedStringsIndexNew.end(); ++it)
	{
		it->second = nSharedStringsIndexNew++;
	}
	nSstCount = 0;
	for (n32 i = 0; i < static_cast<n32>(vSheetName.size()); i++)
	{
		const wstring& sTableName = vSheetName[i];
		SSheetInfo& sheetInfo = mSheetInfo[sTableName];
		map<n32, n32>& mRowStyle = mTableRowStyle[sTableName];
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
		for (map<n32, map<n32, pair<bool, wstring>>>::iterator itRow = mRowColumnText.begin(); itRow != mRowColumnText.end(); ++itRow)
		{
			n32 nRowIndex = itRow->first;
			map<n32, pair<bool, wstring>>& mColumnText = itRow->second;
			n32 nRowStyle = mRowStyle[nRowIndex];
			sSheetXml += Format("<row r=\"%d\" spans=\"1:%d\" s=\"%d\" customFormat=\"1\" x14ac:dyDescent=\"0.2\">", nRowIndex + 1, sheetInfo.ColumnCount, nRowStyle);
			for (map<n32, pair<bool, wstring>>::iterator itColumn = mColumnText.begin(); itColumn != mColumnText.end(); ++itColumn)
			{
				n32 nColumnIndex = itColumn->first;
				pair<bool, wstring>& pText = itColumn->second;
				if (!pText.first)
				{
					continue;
				}
				nSstCount++;
				wstring& sText = pText.second;
				sSheetXml += Format("<c r=\"%s\" s=\"%d\" t=\"s\">", encodePosition(nRowIndex, nColumnIndex).c_str(), nRowStyle);
				sSheetXml += Format("<v>%d</v>", mSharedStringsIndexNew[sText]);
				sSheetXml += "</c>";
			}
			sSheetXml += "</row>";
		}
		sSheetXml += "</sheetData>";
		sSheetXml += "<phoneticPr fontId=\"1\" type=\"noConversion\"/>";
		sSheetXml += "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>";
		sSheetXml += "<pageSetup paperSize=\"9\" orientation=\"portrait\" r:id=\"rId1\"/>";
		sSheetXml += "</worksheet>";
		UString sSheetXmlPath = a_sXlsxDirName + Format(USTR("/xl/worksheets/sheet%d.xml"), i + 1);
		fp = UFopen(sSheetXmlPath.c_str(), USTR("wb"), false);
		if (fp == nullptr)
		{
			return 1;
		}
		fwrite(sSheetXml.c_str(), 1, sSheetXml.size(), fp);
		fclose(fp);
	}
#pragma endregion /xl/worksheets/sheet%d.xml

#pragma region write /xl/sharedStrings.xml
	// write /xl/sharedStrings.xml
	string sSharedStringsXml;
	sSharedStringsXml += "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
	sSharedStringsXml += Format("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"%d\" uniqueCount=\"%d\">", nSstCount, static_cast<n32>(mSharedStringsIndexNew.size()));
	for (map<wstring, n32>::iterator it = mSharedStringsIndexNew.begin(); it != mSharedStringsIndexNew.end(); ++it)
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
	fp = UFopen(sSharedStringsXmlPath.c_str(), USTR("wb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	fwrite(sSharedStringsXml.c_str(), 1, sSharedStringsXml.size(), fp);
	fclose(fp);
#pragma endregion write /xl/sharedStrings.xml

#pragma region write /xl/styles.xml
	// write /xl/styles.xml
	UString sSwitchGamesStylesXmlPath = sModuleDirName + USTR("/Switch Games/xl/styles.xml");
	if (!copyFile(sStylesXmlPath, sSwitchGamesStylesXmlPath))
	{
		return 1;
	}
#pragma endregion write /xl/styles.xml

#pragma region write /xl/workbook.xml
	// write /xl/workbook.xml
	UString sSwitchGamesWorkbookXmlPath = sModuleDirName + USTR("/Switch Games/xl/workbook.xml");
	fp = UFopen(sSwitchGamesWorkbookXmlPath.c_str(), USTR("rb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	eError = workbookXmlDoc.LoadFile(fp);
	fclose(fp);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		return 1;
	}
	pWorkbookXmlDocWorkbook = workbookXmlDoc.RootElement();
	pWorkbookXmlDocWorkbookBookViews = pWorkbookXmlDocWorkbook->FirstChildElement("bookViews");
	if (pWorkbookXmlDocWorkbookBookViews != nullptr)
	{
		tinyxml2::XMLElement* pBookViewsWorkbookView = pWorkbookXmlDocWorkbookBookViews->FirstChildElement("workbookView");
		if (pBookViewsWorkbookView != nullptr)
		{
			if (nActiveTabNew == 0)
			{
				pBookViewsWorkbookView->DeleteAttribute("activeTab");
			}
			else
			{
				pBookViewsWorkbookView->SetAttribute("activeTab", nActiveTabNew);
			}
		}
	}
	tinyxml2::XMLElement* pWorkbookXmlDocWorkbookSheets = pWorkbookXmlDocWorkbook->FirstChildElement("sheets");
	pWorkbookXmlDocWorkbookSheets->DeleteChildren();
	for (n32 i = 0; i < static_cast<n32>(vSheetName.size()); i++)
	{
		tinyxml2::XMLElement* pSheetsSheet = pWorkbookXmlDocWorkbookSheets->InsertNewChildElement("sheet");
		pSheetsSheet->SetAttribute("name", WToU8(vSheetName[i]).c_str());
		pSheetsSheet->SetAttribute("sheetId", i + 1);
		pSheetsSheet->SetAttribute("r:id", Format("rId%d", i + 1).c_str());
	}
	fp = UFopen(sWorkbookXmlPath.c_str(), USTR("wb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	eError = workbookXmlDoc.SaveFile(fp, bCompact);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		fclose(fp);
		return 1;
	}
	fclose(fp);
#pragma endregion write /xl/workbook.xml

#pragma region /[Content_Types].xml
	// /[Content_Types].xml
	UString s_Content_Types_XmlPath = a_sXlsxDirName + USTR("/[Content_Types].xml");
	tinyxml2::XMLDocument _Content_Types_XmlDoc;
	UString sSwitchGames_Content_Types_XmlPath = sModuleDirName + USTR("/Switch Games/[Content_Types].xml");
	fp = UFopen(sSwitchGames_Content_Types_XmlPath.c_str(), USTR("rb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	eError = _Content_Types_XmlDoc.LoadFile(fp);
	fclose(fp);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		return 1;
	}
	tinyxml2::XMLElement* p_Content_Types_XmlDocTypes = _Content_Types_XmlDoc.RootElement();
	p_Content_Types_XmlDocTypes->DeleteChildren();
	tinyxml2::XMLElement* p_Content_Types_XmlDocTypesDefault = p_Content_Types_XmlDocTypes->InsertNewChildElement("Default");
	p_Content_Types_XmlDocTypesDefault->SetAttribute("Extension", "rels");
	p_Content_Types_XmlDocTypesDefault->SetAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
	p_Content_Types_XmlDocTypesDefault = p_Content_Types_XmlDocTypes->InsertNewChildElement("Default");
	p_Content_Types_XmlDocTypesDefault->SetAttribute("Extension", "xml");
	p_Content_Types_XmlDocTypesDefault->SetAttribute("ContentType", "application/xml");
	tinyxml2::XMLElement* p_Content_Types_XmlDocTypesOverride = p_Content_Types_XmlDocTypes->InsertNewChildElement("Override");
	p_Content_Types_XmlDocTypesOverride->SetAttribute("PartName", "/xl/workbook.xml");
	p_Content_Types_XmlDocTypesOverride->SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
	for (n32 i = 0; i < static_cast<n32>(vSheetName.size()); i++)
	{
		p_Content_Types_XmlDocTypesOverride = p_Content_Types_XmlDocTypes->InsertNewChildElement("Override");
		p_Content_Types_XmlDocTypesOverride->SetAttribute("PartName", Format("/xl/worksheets/sheet%d.xml", i + 1).c_str());
		p_Content_Types_XmlDocTypesOverride->SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
	}
	p_Content_Types_XmlDocTypesOverride = p_Content_Types_XmlDocTypes->InsertNewChildElement("Override");
	p_Content_Types_XmlDocTypesOverride->SetAttribute("PartName", "/xl/theme/theme1.xml");
	p_Content_Types_XmlDocTypesOverride->SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.theme+xml");
	p_Content_Types_XmlDocTypesOverride = p_Content_Types_XmlDocTypes->InsertNewChildElement("Override");
	p_Content_Types_XmlDocTypesOverride->SetAttribute("PartName", "/xl/styles.xml");
	p_Content_Types_XmlDocTypesOverride->SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
	p_Content_Types_XmlDocTypesOverride = p_Content_Types_XmlDocTypes->InsertNewChildElement("Override");
	p_Content_Types_XmlDocTypesOverride->SetAttribute("PartName", "/xl/sharedStrings.xml");
	p_Content_Types_XmlDocTypesOverride->SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
	p_Content_Types_XmlDocTypesOverride = p_Content_Types_XmlDocTypes->InsertNewChildElement("Override");
	p_Content_Types_XmlDocTypesOverride->SetAttribute("PartName", "/docProps/core.xml");
	p_Content_Types_XmlDocTypesOverride->SetAttribute("ContentType", "application/vnd.openxmlformats-package.core-properties+xml");
	p_Content_Types_XmlDocTypesOverride = p_Content_Types_XmlDocTypes->InsertNewChildElement("Override");
	p_Content_Types_XmlDocTypesOverride->SetAttribute("PartName", "/docProps/app.xml");
	p_Content_Types_XmlDocTypesOverride->SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
	fp = UFopen(s_Content_Types_XmlPath.c_str(), USTR("wb"), false);
	if (fp == nullptr)
	{
		return 1;
	}
	eError = _Content_Types_XmlDoc.SaveFile(fp, bCompact);
	if (eError != tinyxml2::XML_SUCCESS)
	{
		fclose(fp);
		return 1;
	}
	fclose(fp);
#pragma endregion /[Content_Types].xml

	return 0;
}

int UMain(int argc, UChar* argv[])
{
	if (argc < 3)
	{
		return 1;
	}
	if (UCslen(argv[1]) == 1)
	{
		switch (*argv[1])
		{
		case USTR('R'):
		case USTR('r'):
			return resaveXlsx(argv[2]);
		default:
			break;
		}
	}
	return 1;
}
