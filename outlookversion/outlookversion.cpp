// outlookversion
// Returns 2013, 2010, 2007, 2003 or 0 if no Outlook version is installed.
//

#include <Windows.h>
#include <Msi.h>

#include "stdafx.h"

int getOutlookVersion(void)
{
    TCHAR outlookRegister[][MAX_PATH] = {
        TEXT("{E83B4360-C208-4325-9504-0D23003A74A5}"), // Outlook 2013
        TEXT("{1E77DE88-BCAB-4C37-B9E5-073AF52DFD7A}"), // Outlook 2010
        TEXT("{24AAE126-0911-478F-A019-07B875EB9996}"), // Outlook 2007
        TEXT("{BC174BAD-2F53-4855-A1D5-0D575C19B1EA}")  // Outlook 2003
    };

    int outlookVersions[] = {
        2013, // Outlook 2013
        2010, // Outlook 2010
        2007, // Outlook 2007
        2003 // Outlook 2003
    };
    DWORD pathLength = 0;

    for(int i = 0; i < (sizeof(outlookVersions) / sizeof(outlookVersions[0])); ++i)
    {
        if(MsiProvideQualifiedComponent(
                    outlookRegister[i],
                    TEXT("outlook.x64.exe"),
                    (DWORD) INSTALLMODE_DEFAULT,
                    NULL,
                    &pathLength)
                == ERROR_SUCCESS)
        {
            return outlookVersions[i];
        }
        else if(MsiProvideQualifiedComponent(
                    outlookRegister[i],
                    TEXT("outlook.exe"),
                    (DWORD) INSTALLMODE_DEFAULT,
                    NULL,
                    &pathLength)
                == ERROR_SUCCESS)
        {
            return outlookVersions[i];
        }
    }

    return 0;
}

int _tmain(int argc, _TCHAR* argv[])
{
    BOOL is64bit = FALSE;
    LPSTR version[MAX_PATH * sizeof(TCHAR)];
    HRESULT result;

    printf("%d %d\n", getOutlookVersion(), INSTALLMODE_DEFAULT);
    
    return 0;
}

