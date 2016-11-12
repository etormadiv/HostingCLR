// HostingCLR.cpp : définit le point d'entrée pour l'application console.
//

//By Etor Madiv

//Info: Assemblies that are calling Assembly.GetEntryAssembly() method will fail to load if exceptions are not catched!
//Info: Assembly.GetEntryAssembly() method will return null if loaded by this code, se it should NOT be used
//Info: Please fix any code that use Assembly.GetEntryAssembly() and change it typeof(MyType).Assembly instead
//Info: Please read https://msdn.microsoft.com/en-us/library/system.reflection.assembly.getentryassembly(v=vs.110).aspx#Anchor_1

#include "stdafx.h"

#define RAW_ASSEMBLY_LENGTH 8192

unsigned char rawData[8192] = {
	//...
};


int _tmain(int argc, _TCHAR* argv[])
{
	ICLRMetaHost* pMetaHost = NULL;

	HRESULT hr;

	/* Get ICLRMetaHost instance */

	hr = CLRCreateInstance(CLSID_CLRMetaHost, IID_ICLRMetaHost, (VOID**)&pMetaHost);

	if(FAILED(hr))
	{
		printf("[!] CLRCreateInstance(...) failed\n");

		getchar();

		return -1;
	}

	printf("[+] CLRCreateInstance(...) succeeded\n");

	ICLRRuntimeInfo* pRuntimeInfo = NULL;

	/* Get ICLRRuntimeInfo instance */

	hr = pMetaHost->GetRuntime(L"v2.0.50727", IID_ICLRRuntimeInfo, (VOID**)&pRuntimeInfo);

	if(FAILED(hr))
	{
		printf("[!] pMetaHost->GetRuntime(...) failed\n");

		getchar();

		return -1;
	}

	printf("[+] pMetaHost->GetRuntime(...) succeeded\n");

	BOOL bLoadable;

	/* Check if the specified runtime can be loaded */

	hr = pRuntimeInfo->IsLoadable(&bLoadable);

	if(FAILED(hr) || !bLoadable)
	{
		printf("[!] pRuntimeInfo->IsLoadable(...) failed\n");

		getchar();

		return -1;
	}

	printf("[+] pRuntimeInfo->IsLoadable(...) succeeded\n");
	
	ICorRuntimeHost* pRuntimeHost = NULL;

	/* Get ICorRuntimeHost instance */

	hr = pRuntimeInfo->GetInterface(CLSID_CorRuntimeHost, IID_ICorRuntimeHost, (VOID**)&pRuntimeHost);

	if(FAILED(hr))
	{
		printf("[!] pRuntimeInfo->GetInterface(...) failed\n");

		getchar();

		return -1;
	}

	printf("[+] pRuntimeInfo->GetInterface(...) succeeded\n");
	
	/* Start the CLR */

	hr = pRuntimeHost->Start();

	if(FAILED(hr))
	{
		printf("[!] pRuntimeHost->Start() failed\n");

		getchar();

		return -1;
	}

	printf("[+] pRuntimeHost->Start() succeeded\n");
	
	IUnknownPtr pAppDomainThunk = NULL;

	hr = pRuntimeHost->GetDefaultDomain(&pAppDomainThunk);

	if(FAILED(hr))
	{
		printf("[!] pRuntimeHost->GetDefaultDomain(...) failed\n");

		getchar();

		return -1;
	}

	printf("[+] pRuntimeHost->GetDefaultDomain(...) succeeded\n");
	
	_AppDomainPtr pDefaultAppDomain = NULL;

	/* Equivalent of System.AppDomain.CurrentDomain in C# */

	hr = pAppDomainThunk->QueryInterface(__uuidof(_AppDomain), (VOID**) &pDefaultAppDomain);

	if(FAILED(hr))
	{
		printf("[!] pAppDomainThunk->QueryInterface(...) failed\n");

		getchar();

		return -1;
	}

	printf("[+] pAppDomainThunk->QueryInterface(...) succeeded\n");

	_AssemblyPtr pAssembly = NULL;

	SAFEARRAYBOUND rgsabound[1];

	rgsabound[0].cElements = RAW_ASSEMBLY_LENGTH;

	rgsabound[0].lLbound   = 0;

	SAFEARRAY* pSafeArray  = SafeArrayCreate(VT_UI1, 1, rgsabound);

	void* pvData = NULL;

	hr = SafeArrayAccessData(pSafeArray, &pvData);

	if(FAILED(hr))
	{
		printf("[!] SafeArrayAccessData(...) failed\n");

		getchar();

		return -1;
	}

	printf("[+] SafeArrayAccessData(...) succeeded\n");

	memcpy(pvData, rawData, RAW_ASSEMBLY_LENGTH);

	hr = SafeArrayUnaccessData(pSafeArray);

	if(FAILED(hr))
	{
		printf("[!] SafeArrayUnaccessData(...) failed\n");

		getchar();

		return -1;
	}

	printf("[+] SafeArrayUnaccessData(...) succeeded\n");

	/* Equivalent of System.AppDomain.CurrentDomain.Load(byte[] rawAssembly) */

	hr = pDefaultAppDomain->Load_3(pSafeArray, &pAssembly);

	if(FAILED(hr))
	{
		printf("[!] pDefaultAppDomain->Load_3(...) failed\n");

		getchar();

		return -1;
	}

	printf("[+] pDefaultAppDomain->Load_3(...) succeeded\n");

	_MethodInfoPtr pMethodInfo = NULL;

	/* Assembly.EntryPoint Property */

	hr = pAssembly->get_EntryPoint(&pMethodInfo);

	if(FAILED(hr))
	{
		printf("[!] pAssembly->get_EntryPoint(...) failed\n");

		getchar();

		return -1;
	}

	printf("[+] pAssembly->get_EntryPoint(...) succeeded\n");

	VARIANT retVal;
	ZeroMemory(&retVal, sizeof(VARIANT));

	VARIANT obj;
	ZeroMemory(&obj, sizeof(VARIANT));
	obj.vt = VT_NULL;

	//TODO! Change cElement to the number of Main arguments
	SAFEARRAY *psaStaticMethodArgs = SafeArrayCreateVector(VT_VARIANT, 0, 0);

	/* EntryPoint.Invoke(null, new object[0]) */

	hr = pMethodInfo->Invoke_3(obj, psaStaticMethodArgs, &retVal);

	if(FAILED(hr))
	{
		printf("[!] pMethodInfo->Invoke_3(...) failed, hr = %X\n", hr);

		getchar();

		return -1;
	}

	printf("[+] pMethodInfo->Invoke_3(...) succeeded\n");

	getchar();

	return 0;
}

