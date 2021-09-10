#include "SwitchGamesXlsx.h"
#include <tinyxml2.h>

int resaveXlsx(const UString& a_sXlsxDirName)
{
	CSwitchGamesXlsx switchGamesXlsx;
	switchGamesXlsx.SetXlsxDirName(a_sXlsxDirName);
	return switchGamesXlsx.Resave();
}

int exportXlsx(const UString& a_sXlsxDirName, const UString& a_sListDirName)
{
	if (resaveXlsx(a_sXlsxDirName) != 0)
	{
		return 1;
	}

	CSwitchGamesXlsx switchGamesXlsx;
	switchGamesXlsx.SetXlsxDirName(a_sXlsxDirName);
	switchGamesXlsx.SetListDirName(a_sListDirName);
	return switchGamesXlsx.Export();
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
		default:
			break;
		}
	}
	return 1;
}
