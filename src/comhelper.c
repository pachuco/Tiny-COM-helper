#include <windows.h>
#include "comhelper.h"

void chelp_GUID2strA(LPSTR buf, const GUID* ri) {
    if (ri) {
        wsprintfA(buf, "{%08lX-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}",
            ri->Data1, ri->Data2, ri->Data3, ri->Data4[0],
            ri->Data4[1], ri->Data4[2], ri->Data4[3],
            ri->Data4[4], ri->Data4[5], ri->Data4[6],
            ri->Data4[7]);
    } else {
        wsprintfA(buf, "{--------------NULL PTR--------------}"); //string must be of same size
    }
}

void chelp_GUID2strW(LPWSTR buf, const GUID* ri) {
    LPSTR bufA = (LPSTR)buf;
    CHAR guidStr[GUIDSTR_SIZE+1];
    chelp_GUID2strA(guidStr, ri);
    
    for (UINT i=0; i < GUIDSTR_SIZE+1; i++) {
        *bufA++ = guidStr[i];
        *bufA++ = '0';
    }
}

BOOL chelp_cmpMultGUID(const GUID* p1, const GUID** pp2, int count) {
    for (int i=0; i < count; i++) {
        if (IsEqualGUID(p1, pp2[i])) return TRUE;
    }
    return FALSE;
}