#ifndef SWITCH_GAMES_XLSX_H_
#define SWITCH_GAMES_XLSX_H_

#include <sdw.h>

namespace tinyxml2
{
	class XMLElement;
}

class CSwitchGamesXlsx
{
public:
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
	CSwitchGamesXlsx();
	~CSwitchGamesXlsx();
	void SetXlsxDirName(const UString& a_sXlsxDirName);
	void SetTableDirName(const UString& a_sTableDirName);
	int Resave();
	int Export();
	int Sort();
private:
	static string validateText(const string& a_sText);
	static string validateValue(const string& a_sValue);
	static bool travelElement(const tinyxml2::XMLElement* a_pRootElement, string& a_sXml);
	static string encodePosition(n32 a_nRowIndex, n32 a_nColumnIndex);
	static bool copyFile(const UString& a_sDestFileName, const UString& a_sSrcFileName);
	static string trim(const string& a_sLine);
	static bool empty(const string& a_sLine);
	static bool rowColumnTextCompare(const pair<n32, wstring>& lhs, const pair<n32, wstring>& rhs);
	int readConfig();
	int readWorkbook();
	int readSharedStrings();
	int readStyles();
	int resaveRels();
	int resaveApp();
	int resaveCore();
	int resaveWorkbookRels();
	int resaveTheme1();
	int readSheet();
	int writeSheet();
	int writeSharedStrings();
	int resaveStyles();
	int resaveWorkbook();
	int resaveContentTypes();
	int readTable();
	int writeTable();
	int sortTable();
	UString m_sXlsxDirName;
	UString m_sTableDirName;
	bool m_bResave;
	bool m_bCompact;
	UString m_sModuleDirName;
	vector<wstring> m_vSheetName;
	set<wstring> m_sSheetName;
	map<wstring, n32> m_mSheetIndexOld;
	string m_sCreator;
	string m_sLastModifiedBy;
	string m_sLastModifiedDateTime;
	n32 m_nActiveTabOld;
	n32 m_nActiveTabNew;
	n32 m_nSstCount;
	map<n32, wstring> m_mSharedStrings;
	map<wstring, n32> m_mSharedStringsIndexNew;
	map<n32, n32> m_mStyleOldToNew;
	map<wstring, SSheetInfo> mSheetInfo;
	map<wstring, map<n32, n32>> mTableRowStyle;
	map<wstring, map<n32, map<n32, pair<bool, wstring>>>> mTableRowColumnText;
};

#endif	// SWITCH_GAMES_XLSX_H_
