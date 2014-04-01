// LaunchMSIUnelevated.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"

#include <iostream>

// use the shell view for the desktop using the shell windows automation to find the
// desktop web browser and then grabs its view
//
// returns:
//      IShellView, IFolderView and related interfaces

HRESULT GetShellViewForDesktop(REFIID riid, void **ppv)
{
	*ppv = NULL;

	IShellWindows *psw;
	HRESULT hr = CoCreateInstance(CLSID_ShellWindows, NULL, CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&psw));
	if (SUCCEEDED(hr))
	{
		HWND hwnd;
		IDispatch* pdisp;
		VARIANT vEmpty = {}; // VT_EMPTY
		if (S_OK == psw->FindWindowSW(&vEmpty, &vEmpty, SWC_DESKTOP, (long*)&hwnd, SWFO_NEEDDISPATCH, &pdisp))
		{
			IShellBrowser *psb;
			hr = IUnknown_QueryService(pdisp, SID_STopLevelBrowser, IID_PPV_ARGS(&psb));
			if (SUCCEEDED(hr))
			{
				IShellView *psv;
				hr = psb->QueryActiveShellView(&psv);
				if (SUCCEEDED(hr))
				{
					hr = psv->QueryInterface(riid, ppv);
					psv->Release();
				}
				psb->Release();
			}
			pdisp->Release();
		}
		else
		{
			hr = E_FAIL;
		}
		psw->Release();
	}
	return hr;
}

// From a shell view object gets its automation interface and from that gets the shell
// application object that implements IShellDispatch2 and related interfaces.
HRESULT GetShellDispatchFromView(IShellView *psv, REFIID riid, void **ppv)
{
	*ppv = NULL;

	IDispatch *pdispBackground;
	HRESULT hr = psv->GetItemObject(SVGIO_BACKGROUND, IID_PPV_ARGS(&pdispBackground));
	if (SUCCEEDED(hr))
	{
		IShellFolderViewDual *psfvd;
		hr = pdispBackground->QueryInterface(IID_PPV_ARGS(&psfvd));
		if (SUCCEEDED(hr))
		{
			IDispatch *pdisp;
			hr = psfvd->get_Application(&pdisp);
			if (SUCCEEDED(hr))
			{
				hr = pdisp->QueryInterface(riid, ppv);
				pdisp->Release();
			}
			psfvd->Release();
		}
		pdispBackground->Release();
	}
	return hr;
}


HRESULT ShellExecInExplorerProcess(PCWSTR pszExeFile, PCWSTR pszParams)
{
	IShellView *psv;
	HRESULT hr = GetShellViewForDesktop(IID_PPV_ARGS(&psv));
	if (SUCCEEDED(hr))
	{
		IShellDispatch2 *psd;
		hr = GetShellDispatchFromView(psv, IID_PPV_ARGS(&psd));
		if (SUCCEEDED(hr))
		{
			BSTR bstrExeFile = SysAllocString(pszExeFile);
			BSTR bstrParams = SysAllocString(pszParams);
			hr = bstrExeFile ? S_OK : E_OUTOFMEMORY;
			if (SUCCEEDED(hr))
			{
				VARIANT vtEmpty = {}; // VT_EMPTY

				VARIANT varArgs;
				VariantInit(&varArgs);
				varArgs.vt = VT_BSTR;
				varArgs.bstrVal = bstrParams;

				hr = psd->ShellExecuteW(bstrExeFile, varArgs, vtEmpty, vtEmpty, vtEmpty);

				SysFreeString(bstrExeFile);
				SysFreeString(bstrParams);
			}
			psd->Release();
		}
		psv->Release();
	}
	return hr;
}


LPTSTR GetProcessName(DWORD dwProcessID)
{
	HANDLE hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, dwProcessID);
	LPTSTR ptcReturn = NULL;

	// Get the process name.
	if (NULL != hProcess)
	{
		TCHAR szProcessName[MAX_PATH] = TEXT("<unknown>");
		DWORD dwSize = MAX_PATH;

		if (QueryFullProcessImageName(hProcess, 0, szProcessName, &dwSize))
		{
			ptcReturn = new TCHAR[_tcslen(szProcessName) + 1];
			ptcReturn[0] = _T('\0');
			_tcscpy_s(ptcReturn, _tcslen(szProcessName) + 1, (LPCTSTR)szProcessName);
		}

		// Release the handle to the process.
		CloseHandle(hProcess);
	}
	
	return ptcReturn;
}

//returns an array of process Ids for any running msiexec processes
DWORD* GetRunningMSIExecProcessIds()
{
	DWORD aProcesses[4096], cbNeeded, cProcesses;
	LPDWORD aMSIProcessIds = new DWORD[32];
	int iCurrentIndex = 0;
	unsigned int i;

	aMSIProcessIds[0] = 0;

	if (!EnumProcesses(aProcesses, sizeof(aProcesses), &cbNeeded))
	{
		return aMSIProcessIds;
	}
	else
	{
		cProcesses = cbNeeded / sizeof(DWORD);

		// Print the name and process identifier for each process.
		for (i = 0; i < cProcesses; i++)
		{
			if (aProcesses[i] != 0)
			{
				LPTSTR ptcName = GetProcessName(aProcesses[i]);

				if (ptcName != NULL)
				{
					//std::wstring strProcessName = ptcName;
					std::basic_string<TCHAR> link = ptcName;
					std::size_t found = link.find(L"msiexec.exe");

					if (found != std::string::npos)
					{
						aMSIProcessIds[iCurrentIndex++] = aProcesses[i];
						aMSIProcessIds[iCurrentIndex] = 0;
					}

					delete[] ptcName;
					ptcName = NULL;
				}
			}
		}
	}

	return aMSIProcessIds;
}

bool IsProcessRunning(DWORD dwProcessId)
{
	DWORD aProcesses[4096], cbNeeded, cProcesses;
	int iCurrentIndex = 0;
	unsigned int i;
	bool bIsRunning = false;
	
	if (!EnumProcesses(aProcesses, sizeof(aProcesses), &cbNeeded))
	{
		return false;
	}
	else
	{
		cProcesses = cbNeeded / sizeof(DWORD);

		// Print the name and process identifier for each process.
		for (i = 0; i < cProcesses; i++)
		{
			if (aProcesses[i] != 0 && aProcesses[i] == dwProcessId)
			{
				bIsRunning = true;
			}
		}
	}

	return bIsRunning;
}

bool DoesRegKeyExist(HKEY hKeyParent, LPCTSTR szKeyName)
{
	LONG	lResult = ERROR_SUCCESS;
	HKEY    hOpenKey = NULL;
	bool	bExists = false;

	lResult = RegOpenKeyEx(hKeyParent, szKeyName, 0, KEY_READ, &hOpenKey);

	if (lResult == ERROR_SUCCESS)
	{
		bExists = true;		
		RegCloseKey(hOpenKey);
	}

	return bExists;
}

//returns 0 if MSI install completes, 1 otherwise
int WaitForMSIToFinish(DWORD dwMSIExecProcessId)
{
	bool	bContinue = true;
	bool	bHasInstallStarted = false;
	bool	bExists = false;
	int		numWaits = 0;
	int		result = 1;

	printf("Enter WaitForMSIToFinish");

	while (bContinue)
	{
		//first check if process is still running
		if (dwMSIExecProcessId > 0)
		{
			if (IsProcessRunning(dwMSIExecProcessId) == false)
			{
				bContinue = false;
				printf("MSIExec is no longer running...existing\r\n");
				result = 1;

				if (bHasInstallStarted == true)
				{
					result = 0;					
				}
				
				break;
			}
		}
		bExists = DoesRegKeyExist(HKEY_LOCAL_MACHINE, _T("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\InProgress"));

		printf("DoesRegKeyExist returns %d\r\n", bExists);

		if (bHasInstallStarted == false)
		{
			if (bExists == true)
			{
				printf("Install has started!\r\n");
				numWaits = 0;
				bHasInstallStarted = true;
			}
		}
		else
		{
			if (bExists == false)
			{
				bContinue = false;

				printf("Install has completed!\r\n");

				result = 0;

				//final sleep
				Sleep(5000);
			}
		}

		if (bContinue == true)
		{
			if  (bHasInstallStarted == false)
			{
				if (numWaits > 60)
				{
					bContinue = false;
				}				
			}
			else
			{
				if (numWaits > 120)
				{
					bContinue = false;
				}
			}

			++numWaits;
			Sleep(1000);
		}
		
	}	

	printf("Existing WaitForMSIToFinish: %d\r\n", result);

	return result;
}


int _tmain(int argc, _TCHAR* argv[])
{
	int iResult = 1;
	HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
	if (SUCCEEDED(hr))
	{
		LPWSTR *szArgList;
		int argCount;

		szArgList = CommandLineToArgvW(GetCommandLine(), &argCount);

		if (argc == 2)
		{
			LPDWORD aMSIProcessesBefore = GetRunningMSIExecProcessIds();

			ShellExecInExplorerProcess(_T("C:\\windows\\system32\\msiexec.exe"), argv[1]);

			LPDWORD aMSIProcessesAfter = GetRunningMSIExecProcessIds();

			//find the process ID of the new msiexec process
			DWORD dwProcessID = 0;
			DWORD dwCurrentAfter = 0;

			for (int iAfterLoop = 0; iAfterLoop < 32; iAfterLoop++)
			{
				dwCurrentAfter = aMSIProcessesAfter[iAfterLoop];

				if (dwCurrentAfter == 0)
				{
					break;
				}
				else
				{
					bool bFound = false;
					for (int iBeforeLoop = 0; iBeforeLoop < 32; iBeforeLoop++)
					{
						if (aMSIProcessesBefore[iBeforeLoop] == 0)
						{
							break;
						}
						else if (dwCurrentAfter == aMSIProcessesBefore[iBeforeLoop])
						{
							bFound = true;
							break;
						}
					}

					if (bFound == false)
					{
						dwProcessID = dwCurrentAfter;
						break;
					}
				}			
			}

			delete[] aMSIProcessesBefore;
			delete[] aMSIProcessesAfter;

			//now, wait for msi to finish						
			iResult = WaitForMSIToFinish(dwProcessID);
		}
		


		CoUninitialize();
	}

	return iResult;
}

