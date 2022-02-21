#include "SwitchGamesXlsx.h"

int resaveXlsx(const UString& a_sXlsxDirName)
{
	CSwitchGamesXlsx switchGamesXlsx;
	switchGamesXlsx.SetXlsxDirName(a_sXlsxDirName);
	return switchGamesXlsx.Resave();
}

int exportXlsx(const UString& a_sXlsxDirName, const UString& a_sTableDirName)
{
	if (resaveXlsx(a_sXlsxDirName) != 0)
	{
		return 1;
	}

	CSwitchGamesXlsx switchGamesXlsx;
	switchGamesXlsx.SetXlsxDirName(a_sXlsxDirName);
	switchGamesXlsx.SetTableDirName(a_sTableDirName);
	return switchGamesXlsx.Export();
}

int importXlsx(const UString& a_sXlsxDirName, const UString& a_sTableDirName)
{
	if (resaveXlsx(a_sXlsxDirName) != 0)
	{
		return 1;
	}

	CSwitchGamesXlsx switchGamesXlsx;
	switchGamesXlsx.SetXlsxDirName(a_sXlsxDirName);
	switchGamesXlsx.SetTableDirName(a_sTableDirName);
	return switchGamesXlsx.Import();
}

int sortTable(const UString& a_sTableDirName)
{
	CSwitchGamesXlsx switchGamesXlsx;
	switchGamesXlsx.SetTableDirName(a_sTableDirName);
	return switchGamesXlsx.Sort();
}

int checkTable(const UString& a_sTableDirName, const UString& a_sResultFileName, n32 a_nCheckLevel, const UString& a_sCheckFilter)
{
	if (a_nCheckLevel >= CSwitchGamesXlsx::kCheckLevelMax)
	{
		return 1;
	}
	CSwitchGamesXlsx switchGamesXlsx;
	switchGamesXlsx.SetTableDirName(a_sTableDirName);
	switchGamesXlsx.SetResultFileName(a_sResultFileName);
	switchGamesXlsx.SetCheckLevel(a_nCheckLevel);
	switchGamesXlsx.SetCheckFilter(UToU8(a_sCheckFilter));
	return switchGamesXlsx.Check();
}

int makeRclonePatchBat(const UString& a_sTableDirName, const UString& a_sRemoteDirName)
{
	CSwitchGamesXlsx switchGamesXlsx;
	switchGamesXlsx.SetTableDirName(a_sTableDirName);
	switchGamesXlsx.SetRemoteDirName(a_sRemoteDirName);
	return switchGamesXlsx.MakeRclonePatchBat();
}

int makeBaiduPCSGoPatchBat(const UString& a_sTableDirName, const UString& a_sRemoteDirName, const UString& a_sBaiduUserId, bool a_bStyleIsGreen)
{
	CSwitchGamesXlsx switchGamesXlsx;
	switchGamesXlsx.SetTableDirName(a_sTableDirName);
	switchGamesXlsx.SetRemoteDirName(a_sRemoteDirName);
	switchGamesXlsx.SetBaiduUserId(a_sBaiduUserId);
	switchGamesXlsx.SetStyleIsGreen(a_bStyleIsGreen);
	return switchGamesXlsx.MakeBaiduPCSGoPatchBat();
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
			if (argc != 3)
			{
				return 1;
			}
			return resaveXlsx(argv[2]);
		case USTR('E'):
		case USTR('e'):
			if (argc != 4)
			{
				return 1;
			}
			return exportXlsx(argv[2], argv[3]);
		case USTR('I'):
		case USTR('i'):
			if (argc != 4)
			{
				return 1;
			}
			return importXlsx(argv[2], argv[3]);
		case USTR('S'):
		case USTR('s'):
			if (argc != 3)
			{
				return 1;
			}
			return sortTable(argv[2]);
		case USTR('C'):
		case USTR('c'):
			if (argc == 4)
			{
				return checkTable(argv[2], argv[3], CSwitchGamesXlsx::kCheckLevelName, USTR("1-"));
			}
			else if (argc == 5)
			{
				return checkTable(argv[2], argv[3], SToN32(argv[4]), USTR("1-"));
			}
			else if (argc == 6)
			{
				return checkTable(argv[2], argv[3], SToN32(argv[4]), argv[5]);
			}
			else
			{
				return 1;
			}
		default:
			break;
		}
	}
	else if (UCscmp(argv[1], USTR("make_rclone_patch_bat")) == 0)
	{
		if (argc != 4)
		{
			return 1;
		}
		return makeRclonePatchBat(argv[2], argv[3]);
	}
	else if (UCscmp(argv[1], USTR("make_baidupcs-go_patch_bat")) == 0)
	{
		if (argc != 7)
		{
			return 1;
		}
		return makeBaiduPCSGoPatchBat(argv[2], argv[3], argv[4], true) == 0 && makeBaiduPCSGoPatchBat(argv[2], argv[5], argv[6], false) == 0 ? 0 : 1;
	}
	return 1;
}
