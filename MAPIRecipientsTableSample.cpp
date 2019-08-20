// MAPIRecipientsTableSample.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include "MAPIRecipientsTableSample.h"

int main()
{
	HRESULT hResult = S_OK;
	ULONG ulLine = 0;
	
	MAPIINIT_0 mapiInit = {0};				// MAPIINIT structure for initialising MAPI
	LPMAPISESSION pMapiSession = nullptr;	// MAPI session interface pointer
	LPMDB pDefaultStore = nullptr;			// MAPI store interface pointer

	// Initialising MAPI
	hResult = MAPIInitialize(&mapiInit);
	if FAILED(hResult) { ulLine = __LINE__; goto Error; }

	// Logging on to the MAPI subsystem. The three flags will cause for the user to be prompted for the profile to be used
	hResult = MAPILogonEx(0, (LPTSTR)L"", (LPTSTR)L"", MAPI_EXPLICIT_PROFILE | MAPI_LOGON_UI | MAPI_NEW_SESSION, &pMapiSession);
	if FAILED(hResult) { ulLine = __LINE__; goto Error; }

	if (!pMapiSession)
	{
		wprintf(L"Unable to open a MAPI session. \n");
	}

	// Calling OpenDefaultStore to open the default message store and get back a pointer to the store object
	pDefaultStore = OpenDefaultStore(pMapiSession);

	// Calling GetAndOpenCalendarFolder to open the calendar folder and list the entries and recipients
	hResult = GetAndOpenCalendarFolder(pDefaultStore);
	if FAILED(hResult) { ulLine = __LINE__; goto Error; }

	goto CleanUp;
Error:
	std::cout << "FAILED! hr = " << std::hex << hResult << ".  LINE = " << std::dec << ulLine << std::endl;
	goto CleanUp;
CleanUp:
	if (pDefaultStore) pDefaultStore->Release();
	if (pMapiSession)
	{
		hResult = pMapiSession->Logoff(0, 0, 0);
		if FAILED(hResult) { std::cout << "FAILED! hr = " << std::hex << hResult << ".  LINE = " << std::dec << ulLine << std::endl; }
		pMapiSession->Release();
	}
	MAPIUninitialize();
	return 0;
}

// OpenDefaultStore - Opens the default message store in the current session and returns an interface poniter to that store
LPMDB OpenDefaultStore(LPMAPISESSION lpMAPISession)
{
	HRESULT hResult = S_OK;
	ULONG ulLine = 0;

	LPMAPITABLE pStoresTable = nullptr;			// IMAPITable interface pointer for managing the Stores Table in the MAPI Session
	LPSRestriction lpStoreRes = nullptr;		// top level restriction to set up the search for the default store
	LPSRestriction lpStoreResLvl1 = nullptr;	// level 1 restriction for the two conditions : RES_EXIST and RES_PROPERTY
	LPSPropValue lpStorePropVal = nullptr;		// SPropoValue variable for the PR_DEFAULT_STORE property
	LPSRowSet lpStoreRows = nullptr;			// SRowSet pointer that will recieve the rows resulted from our search
	LPMDB lpMdb = nullptr;

	// Setting up an enum and a prop tag array with the props we'll use
	enum { iDisplayName, iEntryId, cptaProps };
	SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_DEFAULT_STORE, PR_ENTRYID };

	// Retrive the message stores table from the session
	hResult = lpMAPISession->GetMsgStoresTable(0, &pStoresTable);
	if FAILED(hResult) { ulLine = __LINE__; goto Error; }

	// Allocate memory for the restriction
	hResult = MAPIAllocateBuffer(sizeof(SRestriction), (LPVOID*)&lpStoreRes);
	if FAILED(hResult) { ulLine = __LINE__; goto Error; }

	hResult = MAPIAllocateMore(sizeof(SRestriction) * 2, (LPVOID)lpStoreRes, (LPVOID*)& lpStoreResLvl1);
	if FAILED(hResult) { ulLine = __LINE__; goto Error; }

	hResult = MAPIAllocateMore(sizeof(SPropValue), (LPVOID)lpStoreRes, (LPVOID*)& lpStorePropVal);
	if FAILED(hResult) { ulLine = __LINE__; goto Error; }

	// Set up restriction to query the profile table
	lpStoreRes->rt = RES_AND;
	lpStoreRes->res.resAnd.cRes = 0x00000002;
	lpStoreRes->res.resAnd.lpRes = lpStoreResLvl1;

	lpStoreResLvl1[0].rt = RES_EXIST;
	lpStoreResLvl1[0].res.resExist.ulPropTag = PR_DEFAULT_STORE;
	lpStoreResLvl1[0].res.resExist.ulReserved1 = 0x00000000;
	lpStoreResLvl1[0].res.resExist.ulReserved2 = 0x00000000;
	lpStoreResLvl1[1].rt = RES_PROPERTY;
	lpStoreResLvl1[1].res.resProperty.relop = RELOP_EQ;
	lpStoreResLvl1[1].res.resProperty.ulPropTag = PR_DEFAULT_STORE;
	lpStoreResLvl1[1].res.resProperty.lpProp = lpStorePropVal;

	lpStorePropVal->ulPropTag = PR_DEFAULT_STORE;
	lpStorePropVal->Value.b = true;

	// Query the table to get the the default store only
	hResult = HrQueryAllRows(pStoresTable,
		(LPSPropTagArray)& sptaProps,
		lpStoreRes,
		NULL,
		0,
		&lpStoreRows);
	if FAILED(hResult) { ulLine = __LINE__; goto Error; }

	if (lpStoreRows && lpStoreRows->cRows)
	{
		// Call OpenMsgStore to open the actual store entry
		hResult = lpMAPISession->OpenMsgStore(0, lpStoreRows->aRow[0].lpProps[iEntryId].Value.bin.cb, reinterpret_cast<LPENTRYID>(lpStoreRows->aRow[0].lpProps[iEntryId].Value.bin.lpb), NULL, MAPI_BEST_ACCESS, &lpMdb);
		if FAILED(hResult) { ulLine = __LINE__; goto Error; }
	}
	goto CleanUp;
Error:
	std::cout << "FAILED! hr = " << std::hex << hResult << ".  LINE = " << std::dec << ulLine << std::endl;
	goto CleanUp;
CleanUp:
	if (lpStoreRows) FreeProws(lpStoreRows);
	MAPIFreeBuffer(lpStoreRes);
	if (pStoresTable) pStoresTable->Release();
	return lpMdb;
}

// OpenCalendarFolder - Opens the calendar folder and calls ListFolderEntries to list the specified number of messages
HRESULT OpenCalendarFolder(LPMDB lpMdb, LPSBinary pSBinary)
{
	HRESULT hResult = S_OK;
	ULONG ulLine = 0;
	ULONG ulObjType = 0;
	LPMAPIFOLDER pCalendarFolder = nullptr;

	if (pSBinary)
	{
		// Open the MAPIFolder corresponding to the given EntryID (pSBinary)
		hResult = lpMdb->OpenEntry(pSBinary->cb, reinterpret_cast<LPENTRYID>(pSBinary->lpb), 0, MAPI_BEST_ACCESS, &ulObjType, reinterpret_cast<LPUNKNOWN*>(&pCalendarFolder));
		if FAILED(hResult) { ulLine = __LINE__; goto Error; }
		// Calling ListFolderEntries to list 10 items and their recipients
		hResult = ListFolderEntries(pCalendarFolder, 10);
		if FAILED(hResult) { ulLine = __LINE__; goto Error; }
	}
	goto CleanUp;
Error:
	std::cout << "FAILED! hr = " << std::hex << hResult << ".  LINE = " << std::dec << ulLine << std::endl;
	goto CleanUp;
CleanUp:
	if (pCalendarFolder) pCalendarFolder->Release();
	return hResult;
}

// GetAndOpenCalendarFolder - Opens the calendar folder
HRESULT GetAndOpenCalendarFolder(LPMDB lpMdb)
{
	HRESULT hResult = S_OK;
	ULONG ulLine = 0;
	LPSPropValue pSPropValue = nullptr;
	ULONG ulObjType = 0;
	SBinary entryId = { 0 };
	LPMAPIFOLDER pInboxFolder = nullptr;

	// The calendar folder EntryID is stored on the Inbox folder, for which reason we're calling GetReceiveFolder to get the EntryID of the inbox folder
	hResult = lpMdb->GetReceiveFolder(const_cast<LPTSTR>(L"IPM.Note"), // this is the class of message we want, where IPM.Note is the default message class for Inbox
		MAPI_UNICODE, // flags
		&entryId.cb, // size and...
		reinterpret_cast<LPENTRYID*>(&entryId.lpb), // value of entry ID
		nullptr); // returns a message class if not NULL)
	if FAILED(hResult) { ulLine = __LINE__; goto Error; }

	if (entryId.lpb)
	{
		// if we have a valid entry ID for the Inbox folder then we go ahead and open the Inbox folder
		hResult = lpMdb->OpenEntry(entryId.cb, reinterpret_cast<LPENTRYID>(entryId.lpb), 0, MAPI_BEST_ACCESS, &ulObjType, reinterpret_cast<LPUNKNOWN*>(&pInboxFolder));
		if FAILED(hResult) { ulLine = __LINE__; goto Error; }
		if (pInboxFolder)
		{
			LPSPropValue pSpropValue = nullptr;
			// We then go ahead and read the PR_IPM_APPOINTMENT_ENTRYID from the Inbox folder, this will give us the Entry ID of the calendar folder
			hResult = HrGetOneProp(pInboxFolder, PR_IPM_APPOINTMENT_ENTRYID, &pSpropValue);
			if FAILED(hResult) { ulLine = __LINE__; goto Error; }

			if (pSpropValue)
			{
				// if we have a valid SPBinary then go ahead and call OpenCalendarFolder to open the folder
				hResult = OpenCalendarFolder(lpMdb, (LPSBinary)&pSpropValue->Value.bin);
				if FAILED(hResult) { ulLine = __LINE__; goto Error; }
			}
		}
	}
	goto CleanUp;
Error:
	std::cout << "FAILED! hr = " << std::hex << hResult << ".  LINE = " << std::dec << ulLine << std::endl;
	goto CleanUp;
CleanUp:
	if (pSPropValue) MAPIFreeBuffer(pSPropValue);
	return hResult;
}

// ListFolderEntries - retrieves the given number of entries and calls OpenEntry to list the Subject and process the recipients
HRESULT ListFolderEntries(LPMAPIFOLDER lpMapiFolder, ULONG numberOfEntries)
{
	HRESULT hResult = S_OK;
	ULONG ulLine = 0;
	LPMAPITABLE pContentsTable = nullptr;
	LPSRowSet pCalendarItemRows = nullptr;
	enum
	{
		iSubject,
		iEntryId,
		numCols
	};
	static const SizedSPropTagArray(numCols, rgColProps) = {
		numCols,
		PR_SUBJECT,
		PR_ENTRYID,
	};
	
	// Get the contents table of the folder
	hResult = lpMapiFolder->GetContentsTable(MAPI_UNICODE, &pContentsTable);
	if FAILED(hResult) { ulLine = __LINE__; goto Error; }

	// Reduce the number of columns returned to the columns of interest only 
	hResult = pContentsTable->SetColumns(LPSPropTagArray(&rgColProps), TBL_ASYNC);
	if FAILED(hResult) { ulLine = __LINE__; goto Error; }

	// Runs a search to retrieve the requested number of entries and the properites selected above
	hResult = pContentsTable->QueryRows(numberOfEntries, NULL, &pCalendarItemRows);
	if FAILED(hResult) { ulLine = __LINE__; goto Error; }
	
	if (!pCalendarItemRows || !pCalendarItemRows->cRows)
	{
		wprintf(L"No calendar entries found.\n");
		hResult = E_FAIL;
		goto Error;
	}

	for (ULONG i = 0; i < pCalendarItemRows->cRows; i++)
	{
		wprintf(L"Opening item with subject \"%ls\" \n", pCalendarItemRows->aRow[i].lpProps[iSubject].Value.lpszW);
		// Open each entry in turn by calling OpenEntry
		hResult = OpenEntry(lpMapiFolder, reinterpret_cast<LPSBinary>(&pCalendarItemRows->aRow[i].lpProps[iEntryId].Value.bin));
		if FAILED(hResult) { ulLine = __LINE__; goto Error; }
		wprintf(L"\n");
	}
	goto CleanUp;
Error:
	std::cout << "FAILED! hr = " << std::hex << hResult << ".  LINE = " << std::dec << ulLine << std::endl;
	goto CleanUp;
CleanUp:
	if (pCalendarItemRows) FreeProws(pCalendarItemRows);
	if (pContentsTable) pContentsTable->Release();
	return hResult;
}

// OpenEntry - Opens the given IMessage object and calls PrintAndUpdateRecipients
HRESULT OpenEntry(LPMAPIFOLDER lpMapiFolder, LPSBinary lpSBinary)
{
	HRESULT hResult = S_OK;
	ULONG ulLine = 0;
	ULONG ulObjType = 0;
	LPMESSAGE pMessage = nullptr;

	// Setting up an enum and a prop tag array with the props we'll use
	enum { iDisplayName, iEntryId, cptaProps };
	SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_DISPLAY_NAME, PR_ENTRYID };

	// Openm the given message entry
	hResult = lpMapiFolder->OpenEntry(lpSBinary->cb, reinterpret_cast<LPENTRYID>(lpSBinary->lpb), 0, MAPI_BEST_ACCESS, &ulObjType, reinterpret_cast<LPUNKNOWN*>(&pMessage));
	if FAILED(hResult) { ulLine = __LINE__; goto Error; }

	if (pMessage)
	{	
		// Call PrintAndUpdateRecipients
		hResult = PrintAndUpdateRecipients(pMessage);
			if FAILED(hResult) { ulLine = __LINE__; goto Error; }
	}
	goto CleanUp;
Error:
	std::cout << "FAILED! hr = " << std::hex << hResult << ".  LINE = " << std::dec << ulLine << std::endl;
	goto CleanUp;
CleanUp:
	if (pMessage) pMessage->Release();
	return hResult;

}

// PrintAndUpdateRecipients - Opens the recipients table, prints the Subject and the properites of each recipient and updates the tracking status to respAccepted if a respNone is detected
HRESULT PrintAndUpdateRecipients(LPMESSAGE lpMessage)
{
	HRESULT hResult = S_OK;
	ULONG ulLine = 0;
	BOOL changesPending = FALSE;
	LPMAPITABLE pRecipientsTable = nullptr;
	LPSRowSet pRecipentRows = nullptr;
	LPADRLIST pAddrList = nullptr;
	enum { iDisplayName, iEmailAddress, iTrackStatus, cptaProps };
	SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_DISPLAY_NAME, PR_EMAIL_ADDRESS, PR_RECIPIENT_TRACKSTATUS };

	// Open the recipients table
	hResult = lpMessage->GetRecipientTable(MAPI_UNICODE, &pRecipientsTable);
	if FAILED(hResult) { ulLine = __LINE__; goto Error; }

	if (pRecipientsTable)
	{
		// Get all rows in the recipients table
		hResult = HrQueryAllRows(
			pRecipientsTable,
			NULL,
			NULL,
			NULL,
			NULL,
			&pRecipentRows
		);
		if FAILED(hResult) { ulLine = __LINE__; goto Error; }

		if (pRecipentRows && pRecipentRows->cRows)
		{
			// SRowSet and ADDRLIST are interchangeable so we use ADDRLIST here instead
			pAddrList = reinterpret_cast<LPADRLIST>(pRecipentRows);
			if (pAddrList->cEntries == 0)
			{
				wprintf(L"The current item has no recipients \n");
			}
			// we walk the recipient collection
			for (unsigned int i = 0; i < pAddrList->cEntries; i++)
			{
				wprintf(L"\tRecipient #%o \n", i);
				// we walk the properties collection for each recipient
				for (unsigned int j = 0; j < pAddrList->aEntries[i].cValues; j++)
				{
					switch (pAddrList->aEntries[i].rgPropVals[j].ulPropTag)
					{
					case PR_DISPLAY_NAME:
						wprintf(L"\t\tDisplay name: \"%ls\"\n", pAddrList->aEntries[i].rgPropVals[j].Value.lpszW);
						break;
					case PR_ADDRTYPE:
						wprintf(L"\t\tAddress type: \"%ls\"\n", pAddrList->aEntries[i].rgPropVals[j].Value.lpszW);
						break;
					case PR_EMAIL_ADDRESS:
						wprintf(L"\t\tAddress: \"%ls\"\n", pAddrList->aEntries[i].rgPropVals[j].Value.lpszW);
						break;
					case PR_RECIPIENT_TRACKSTATUS:
						wprintf(L"\t\tTracking status: \"%o\"\n", pAddrList->aEntries[i].rgPropVals[j].Value.l);
						if (respNone == pAddrList->aEntries[i].rgPropVals[j].Value.l)
						{
							// if respNone then update to respAccepted
							pAddrList->aEntries[i].rgPropVals[j].Value.l = respAccepted;
							changesPending = TRUE;
						}
						break;
					};
				}
			}
			if (changesPending)
			{
				wprintf(L"\tUpdating recipients... \n");
				// if we've modified the recipients then go ahead and commit the changes
				hResult = lpMessage->ModifyRecipients(MODRECIP_MODIFY, pAddrList);
				if FAILED(hResult) { ulLine = __LINE__; goto Error; }
				wprintf(L"\tSaving changes... \n");
				// and save the changes on the IMessage object
				hResult = lpMessage->SaveChanges(KEEP_OPEN_READWRITE);
				if FAILED(hResult) { ulLine = __LINE__; goto Error; }
			}
		}
		else
			wprintf(L"\tThe selected mesasge has no recipients. \n");

	}
	goto CleanUp;
Error:
	std::cout << "FAILED! hr = " << std::hex << hResult << ".  LINE = " << std::dec << ulLine << std::endl;
	goto CleanUp;
CleanUp:

	if (pRecipentRows) FreeProws(pRecipentRows);
	if (pRecipientsTable) pRecipientsTable->Release();
	return hResult;
}

