#include <windows.h>
#include "comhelper.h"

//-------------------------------------------------------------
//COM helper functions

void chelp_GUID2strA(LPSTR buf, GUID* ri) {
    wsprintfA(buf, "{%08lX-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}",
        ri->Data1, ri->Data2, ri->Data3, ri->Data4[0],
        ri->Data4[1], ri->Data4[2], ri->Data4[3],
        ri->Data4[4], ri->Data4[5], ri->Data4[6],
        ri->Data4[7]);
}

void chelp_GUID2strW(LPWSTR buf, GUID* ri) {
    LPSTR bufA = (LPSTR)buf;
    CHAR guidStr[GUIDSTR_SIZE+1];
    chelp_GUID2strA(guidStr, ri);
    
    for (int i=0; i < GUIDSTR_SIZE+1; i++) {
        *bufA++ = guidStr[i];
        *bufA++ = '0';
    }
}

BOOL chelp_cmpMultGUID(GUID* p1, GUID** pp2, int count) {
    for (int i=0; i < count; i++) {
        if (IsEqualGUID(p1, pp2[i])) return TRUE;
    }
    return FALSE;
}

BOOL chelp_unregisterCOM(LPCSTR ownershipMark, REFCLSID pGuid) {
    #define CHK(x) if ((x)!= ERROR_SUCCESS) goto FAIL
    char szGuidPath[64];
    HKEY kRoot=0, kClsid=0;
    
    chelp_GUID2strA(szGuidPath, pGuid);
    CHK(RegOpenKeyExA(HKEY_LOCAL_MACHINE, "Software\\Classes\\CLSID\\", 0, KEY_ALL_ACCESS, &kRoot));
    CHK(RegOpenKeyExA(kRoot, szGuidPath, 0, KEY_ALL_ACCESS, &kClsid));
    //refuse to delete without ownership mark
    CHK(RegQueryValueExA(kClsid, ownershipMark, 0, NULL, NULL, NULL));
    RegDeleteKeyA(kClsid, "InprocServer32");
    RegCloseKey(kClsid);
    RegDeleteKeyA(kRoot, szGuidPath);
    RegCloseKey(kRoot);
    
    return TRUE;
    
    FAIL:
        if (kClsid) RegCloseKey(kClsid);
        if (kRoot)  RegCloseKey(kRoot);
        return FALSE;
    #undef CHK
}

BOOL chelp_registerCOM(LPCSTR ownershipMark, REFCLSID pGuid, LPCSTR pszDllPath, BOOL isExpandable, LPCSTR pszThreadModel, LPCSTR pszDescription) {
    #define CHK(x) if ((x)!= ERROR_SUCCESS) goto FAIL
    char szGuidPath[64] = "Software\\Classes\\CLSID\\";
    HKEY kClsid=0, kInproc=0;
    DWORD type = isExpandable ? REG_EXPAND_SZ : REG_SZ;
    DWORD disp;
    
    chelp_GUID2strA(szGuidPath+23, pGuid);
    //does CLSID already exist?
    if (RegOpenKeyExA(HKEY_LOCAL_MACHINE, szGuidPath, 0, KEY_READ, &kClsid) == ERROR_SUCCESS) {
        //if ownership stamp is not present, back away
        if (RegQueryValueExA(kClsid, ownershipMark, 0, NULL, NULL, NULL) != ERROR_SUCCESS) goto FAIL;
        RegCloseKey(kClsid);
    }
    
    //create CLSID key
    CHK(RegCreateKeyExA(HKEY_LOCAL_MACHINE, szGuidPath, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &kClsid, &disp));
    //write object description
    CHK(RegSetValueExA(kClsid, NULL, 0, REG_SZ, (BYTE*)pszDescription, strlen(pszDescription)));
    //write our ownership stamp
    CHK(RegSetValueExA(kClsid, ownershipMark, 0, REG_NONE, NULL, 0));
    //make inprocserver key
    CHK(RegCreateKeyExA(kClsid, "InprocServer32", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &kInproc, &disp));
    //write path to DLL
    CHK(RegSetValueExA(kInproc, NULL, 0, type, (BYTE*)pszDllPath, strlen(pszDllPath)));
    //write threading model
    CHK(RegSetValueExA(kInproc, "ThreadingModel", 0, REG_SZ, (BYTE*)pszThreadModel, strlen(pszThreadModel)));
    RegCloseKey(kInproc);
    RegCloseKey(kClsid);

    return TRUE;
    
    FAIL:
        if(kInproc) RegCloseKey(kInproc);
        if(kClsid)  RegCloseKey(kClsid);
        chelp_unregisterCOM(ownershipMark, pGuid);
        return FALSE;
    #undef CHK
}

//-------------------------------------------------------------
//COM base helper API

static HRESULT STDMETHODCALLTYPE cbase_FactCreateInstance(COMGenerObj* self, IUnknown *pUnkOuter, REFIID riid, void **ppvObject);
static HRESULT STDMETHODCALLTYPE cbase_FactLockServer(COMGenerObj* self, BOOL flock);

static const IClassFactoryVtbl vt_ClassFactory = {
    &cbase_UnkQueryInterface,
    &cbase_UnkAddRef,
    &cbase_UnkRelease,
    
    &cbase_FactCreateInstance,
    &cbase_FactLockServer
};

static const REFIID factoryIIDList[] = { &IID_IClassFactory, &IID_IUnknown };
static const COMDesc factoryDesc = {
    NULL,
    &factoryIIDList,
    2,
    &vt_ClassFactory,
    sizeof(COMGenerObj),
    NULL, NULL,
    NULL, NULL, NULL
};

static const COMDesc** serverList  = NULL;
static HINSTANCE dllInstance  = NULL;
static int       serverCount  = 0;
static LONG      dllUseCount  = 0;
static LONG      dllLockCount = 0;

static HRESULT err_comNotInit() {
    OutputDebugStringA("COM-base was not initialized prior.");
    return E_FAIL;
}

void cbase_init(HINSTANCE hinst, const COMDesc** cobjArr, int nrObjects) {
    dllInstance = hinst;
    serverList = cobjArr;
    serverCount = nrObjects;
}

HRESULT cbase_createInstance(const COMDesc* conf, void** ppv, BOOL isFactory) {
    HRESULT hr = S_OK;
    const COMDesc* pco = isFactory ? &factoryDesc : conf;
    COMGenerObj* obj;
    
    if(!ppv) return E_POINTER;
    *ppv = NULL;
    obj = GlobalAlloc(GMEM_FIXED | GMEM_ZEROINIT, pco->objSize);
    if (!obj) return E_OUTOFMEMORY;
    
    obj->lpVtbl = pco->rvtbl;
    obj->conf = conf;
    obj->isFactory = isFactory;
    cbase_UnkAddRef(obj);
    if (pco->cbConstruct && !obj->isFactory) hr = pco->cbConstruct(obj, pco->rclsid, NULL);
    if (FAILED(hr)) {
        GlobalFree(obj);
    } else {
        InterlockedIncrement(&dllUseCount);
        *ppv = (void*)obj;
    }
    return hr;
}

//Unknown Object

HRESULT STDMETHODCALLTYPE cbase_UnkQueryInterface(COMGenerObj* self, REFIID riid, void **ppv) {
    const COMDesc* pco = self->isFactory ? &factoryDesc : self->conf;
    if(!ppv) return E_POINTER;
    *ppv = NULL;
    
    if(chelp_cmpMultGUID(riid, pco->riidArr, pco->iidCount)) {
        cbase_UnkAddRef(self);
        *ppv = (void*)self;
        return S_OK;
    }

    return E_NOINTERFACE;
}
ULONG   STDMETHODCALLTYPE cbase_UnkAddRef(COMGenerObj* self) {
    return ++self->count;
}
ULONG   STDMETHODCALLTYPE cbase_UnkRelease(COMGenerObj* self) {
    const COMDesc* pco = self->conf;
    if((--self->count) == 0) {
        if (pco->cbDestruct && !self->isFactory) pco->cbDestruct(self, pco->rclsid, NULL);
        GlobalFree(self);
        InterlockedDecrement(&dllUseCount);
        return 0;
    }
    
    return self->count;
}



//Class Factory Object

static HRESULT STDMETHODCALLTYPE cbase_FactCreateInstance(COMGenerObj* self, IUnknown *pUnkOuter, REFIID riid, void **ppv) {
    const COMDesc* pco = self->conf;
    if(!ppv) return E_POINTER;
    *ppv = NULL;
    if(pUnkOuter) return CLASS_E_NOAGGREGATION; //what do we even do with this?
    
    if (chelp_cmpMultGUID(riid, pco->riidArr, pco->iidCount)) {
        return cbase_createInstance(pco, ppv, FALSE);
    }
    return E_NOINTERFACE;
}

static HRESULT STDMETHODCALLTYPE cbase_FactLockServer(COMGenerObj* self, BOOL flock) {
    flock ? InterlockedIncrement(&dllLockCount) : InterlockedDecrement(&dllLockCount);
    return S_OK;
}



//DLL hooks

HRESULT WINAPI cbase_DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv) {
    if (!dllInstance) return err_comNotInit();
    if(!ppv) return E_POINTER;
    *ppv = NULL;
    
    for (int i=0; i < serverCount; i++) {
        const COMDesc* pco = serverList[i];
        
        if (IsEqualGUID(rclsid, pco->rclsid)) {
            if (chelp_cmpMultGUID(riid, &factoryIIDList, 2)) {
                return cbase_createInstance(pco, ppv, TRUE);
            }
            return E_NOINTERFACE;
        }
    }

    return CLASS_E_CLASSNOTAVAILABLE; 
}

HRESULT WINAPI cbase_DllCanUnloadNow() {
    if (!dllInstance) return err_comNotInit();
    
    return (dllUseCount==0 && dllLockCount==0) ? S_OK : S_FALSE;
}

HRESULT WINAPI cbase_DllUnregisterServer() {
    if (!dllInstance) return err_comNotInit();
    
    for (int i=0; i < serverCount; i++) {
        chelp_unregisterCOM(serverList[i]->ownMark, serverList[i]->rclsid);
    }

    return S_OK;
}

HRESULT WINAPI cbase_DllRegisterServer() {
    if (!dllInstance) return err_comNotInit();
    
    #define CHK(x) if (!(x)) goto FAIL
    char modulePath[MAX_PATH];
    
    
    CHK(GetModuleFileNameA(dllInstance, modulePath, MAX_PATH));
    for (int i=0; i < serverCount; i++) {
        const COMDesc* pco = serverList[i];
        CHK(chelp_registerCOM(pco->ownMark, pco->rclsid, modulePath, FALSE, pco->thModel, pco->descript));
    }

    return S_OK;

    FAIL:
        cbase_DllUnregisterServer();
        return E_FAIL;
    #undef CHK
}