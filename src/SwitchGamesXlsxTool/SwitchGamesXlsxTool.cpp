#include "SwitchGamesXlsx.h"
#include <tinyxml2.h>

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

int sortTable(const UString& a_sTableDirName)
{
	CSwitchGamesXlsx switchGamesXlsx;
	switchGamesXlsx.SetTableDirName(a_sTableDirName);
	return switchGamesXlsx.Sort();
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
		case USTR('S'):
		case USTR('s'):
			if (argc != 3)
			{
				return 1;
			}
			return sortTable(argv[2]);
		default:
			break;
		}
	}
	return 1;
}
