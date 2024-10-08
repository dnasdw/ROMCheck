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
	enum ECheckLevel
	{
		kCheckLevelName = 0,
		kCheckLevelCRLF = 1,
		kCheckLevelSfv = 2,
		kCheckLevelMax
	};
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
	enum EEncoding
	{
		kEncodingUnknown,
		kEncodingCP437,
		kEncodingUTF8,
		kEncodingUTF8withBOM
	};
	enum ELineType
	{
		kLineTypeUnknown,
		kLineTypeLF,
		kLineTypeLF_CR,
		kLineTypeLFMix,
		kLineTypeCRLF,
		kLineTypeCRLF_CR,
		kLineTypeCRLFMix,
		kLineTypeCR,
		kLineTypeCRMix
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
	struct SResult
	{
		n32 Year;
		wstring Name;
		wstring Path;
		bool Exist;
		wstring Type;
		map<UString, u32> RarFile;
		vector<UString> NfoFile;
		vector<UString> SfvFile;
		vector<UString> OtherFile;
	};
	struct STextFileContent
	{
		string TextOld;
		string TextNew;
		EEncoding EncodingOld;
		EEncoding EncodingNew;
		ELineType LineTypeOld;
		ELineType LineTypeNew;
		STextFileContent();
	};
	CSwitchGamesXlsx();
	~CSwitchGamesXlsx();
	void SetXlsxDirName(const UString& a_sXlsxDirName);
	void SetTableDirName(const UString& a_sTableDirName);
	void SetResultFileName(const UString& a_sResultFileName);
	void SetCheckLevel(n32 a_nCheckLevel);
	void SetCheckFilter(const string& a_sCheckFilter);
	void SetRemoteDirName(const UString& a_sRemoteDirName);
	void SetBaiduUserId(const UString& a_sBaiduUserId);
	void SetStyleIsGreen(bool a_bStyleIsGreen);
	int Resave();
	int Export();
	int Import();
	int Sort();
	int Check();
	int MakeRclonePatchBat();
	int MakeBaiduPCSGoPatchBat();
private:
	static string validateText(const string& a_sText);
	static string validateValue(const string& a_sValue);
	static bool travelElement(const tinyxml2::XMLElement* a_pRootElement, string& a_sXml);
	static string encodePosition(n32 a_nRowIndex, n32 a_nColumnIndex);
	static bool copyFile(const UString& a_sDestFileName, const UString& a_sSrcFileName);
	static string trim(const string& a_sLine);
	static wstring trim(const wstring& a_sLine);
	static bool empty(const string& a_sLine);
	static bool pathCompare(const UString& lhs, const UString& rhs);
	static bool rowColumnTextCompare(const pair<n32, wstring>& lhs, const pair<n32, wstring>& rhs);
	static bool fileListCompare(const pair<UString, bool>& lhs, const pair<UString, bool>& rhs);
	static int readTextFile(const UString& a_sFilePath, STextFileContent& a_TextFileContent, bool a_bAllowEmpty);
	static bool makeDir(const UString& a_sDirPath);
	static int writeFileString(const UString& a_sFilePath, const string& a_sStringContent);
	int readConfig();
	int readWorkbook();
	int readSharedStrings();
	int readStyles();
	int resaveRels() const;
	int resaveApp();
	int resaveCore();
	int resaveWorkbookRels() const;
	int resaveTheme1() const;
	int readSheet();
	int writeSheet();
	int writeSharedStrings() const;
	int resaveStyles() const;
	int resaveWorkbook() const;
	int resaveContentTypes() const;
	int readTable();
	int writeTable();
	int sortTable();
	int checkTable();
	int readResult();
	void updateSharedStrings();
	bool matchGreenStyle(const wstring& a_sName) const;
	int makePatchTypeFileList();
	int makeRclonePatchBat() const;
	int makeBaiduPCSGoPatchBat() const;
	UString m_sXlsxDirName;
	UString m_sTableDirName;
	UString m_sResultFileName;
	n32 m_nCheckLevel;
	string m_sCheckFilter;
	UString m_sRemoteDirName;
	UString m_sBaiduUserId;
	bool m_bStyleIsGreen;
	bool m_bResave;
	bool m_bCompact;
	UString m_sModuleDirName;
	vector<wstring> m_vSheetName;
	set<wstring> m_sSheetName;
	vector<wregex> m_vGreenStyleNamePattern;
	vector<wregex> m_vNotGreenStyleNamePattern;
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
	map<wstring, SSheetInfo> m_mSheetInfo;
	map<wstring, map<n32, n32>> m_mTableRowStyle;
	map<wstring, map<n32, map<n32, pair<bool, wstring>>>> m_mTableRowColumnText;
	vector<SResult> m_vResult;
	vector<UString> m_vPatchTypeList;
	vector<pair<UString, bool>> m_vPatchFileList;
};

#endif	// SWITCH_GAMES_XLSX_H_
