#pragma once
#include <iostream>
#include <MAPIX.h>
#include <MapiUtil.h>
#include <Mapidefs.h>
#include <Mapitags.h>
#include <ExtraMAPIDefs.h>
// Entry ID for the Calendar
#define PR_IPM_APPOINTMENT_ENTRYID (PROP_TAG(PT_BINARY, 0x36D0))

#pragma comment(lib, "mapi32.lib")

LPMDB OpenDefaultStore(LPMAPISESSION lpMAPISession);
HRESULT OpenCalendarFolder(LPMDB lpMdb, LPSBinary pSBinary);
HRESULT ListFolderEntries(LPMAPIFOLDER lpMapiFolder, ULONG numberOfEntries);
HRESULT GetAndOpenCalendarFolder(LPMDB lpMdb);
HRESULT ListFolderEntries(LPMAPIFOLDER lpMapiFolder, ULONG numberOfEntries);
HRESULT OpenEntry(LPMAPIFOLDER lpMapiFolder, LPSBinary lpSBinary);
HRESULT PrintAndUpdateRecipients(LPMESSAGE lpMessage);
