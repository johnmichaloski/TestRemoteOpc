



ERROR: DCOM  required impersonation level was not provided
If COM app is administrator will kinda launch, but then die.


Ok if Propertie->Linker-> UAC Execution Level: asInvoker (/level='asInvoker')







============================================================================================================


		Load();

		CoInitialize(NULL);

		// Unnecessary and TOO LATE to take effect in APP
		//ComInitializeSecuity dcom;
		//dcom.dwAuthnLevel=RPC_C_AUTHN_LEVEL_CONNECT;;
		//dcom.dwImpLevel=RPC_C_IMP_LEVEL_IDENTIFY;
		//hr=dcom.InitializeSecurity();


		CLSID gOpcServerClsid;
		// OPC.SINUMERIK.MachineSwitch
		//sClsid=L"{75d00afe-dda5-11d1-b944-9e614d000000}";  // Siemens 840D OPC Server CLSID : 75d00afe-dda5-11d1-b944-9e614d000000
		if(FAILED(CLSIDFromString(sClsid , &gOpcServerClsid)))
		{
			throw bstr_t(L"CLSIDFromString(_bstr_t(sClsid) , &gOpcServerClsid))) FAILED\n");
			return 0;
		}
		TCHAR * domain = L"NIST";

		COAUTHIDENTITY idn; ((LPCWSTR) L".",(LPCWSTR) L"michalos",(LPCWSTR) "jlm%^%^cbm1984");
		SecureZeroMemory(&idn, sizeof(COAUTHIDENTITY));
		idn.Flags = SEC_WINNT_AUTH_IDENTITY_UNICODE;
		idn.User = (USHORT*)L"michalos";
		idn.UserLength = wcslen(L"michalos");
		idn.Password = (USHORT*)L"jlm%^%^cbm1984";
		idn.PasswordLength = wcslen(L"jlm%^%^cbm1984");
		idn.Domain= (USHORT *)  L"NIST";
		idn.DomainLength= wcslen( L"NIST");


		COAUTHINFO AuthInfo;
		AuthInfo.dwAuthnSvc = RPC_C_AUTHN_WINNT;
		AuthInfo.dwAuthzSvc = RPC_C_AUTHZ_NONE;
		AuthInfo.pwszServerPrincName = L"NIST\\mountaineer";
		AuthInfo.dwAuthnLevel = RPC_C_AUTHN_LEVEL_CONNECT;
		AuthInfo.dwImpersonationLevel = RPC_C_IMP_LEVEL_IMPERSONATE;
		AuthInfo.pAuthIdentityData = &idn;
		AuthInfo.dwCapabilities = EOAC_NONE;

		COSERVERINFO srvinfo; // = {0, server, NULL, 0}; // Create the object and query for two interfaces
		srvinfo.dwReserved1=0;
		srvinfo.dwReserved2=0;
		srvinfo.pwszName=server; //L"129.6.72.54"; // L"129.6.78.90";
		srvinfo.pAuthInfo= &AuthInfo; //  &athn;

		MULTI_QI mqi[]=
		{ 
			//{&IID_IUnknown, NULL , 0} 
			{&IID_IOPCServer, NULL , 0} 
		};


		if(FAILED(hr=CoCreateInstanceEx(
			gOpcServerClsid, // Request an instance of class 
			NULL, // No aggregation
			CLSCTX_REMOTE_SERVER, // CLSCTX_SERVER, // Any server is fine
			&srvinfo, // Contains remote server name
			sizeof(mqi)/sizeof(mqi[0]), // number of interfaces we are requesting (2)
			(MULTI_QI        *) &mqi))) // structure indicating IIDs and interface pointers
		{

			throw bstrFormat(L":Test840D CoCreateInstanceEx  FAILED 0x%x = %s\n",  hr, ErrorFormatMessage(hr));
		}
		_pIOPCServer = (IOPCServer *) (mqi[0].pItf); // Retrieve first interface pointer hr=pBackupAdmin->StartBackup(); // use it…

		