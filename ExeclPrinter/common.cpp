#include "stdafx.h"
#include "common.h"

/*
 * AnsiToUnicode converts the ANSI string pszA to a Unicode string
 * and returns the Unicode string through ppszW. Space for the
 * the converted string is allocated by AnsiToUnicode.
 */ 

HRESULT __fastcall AnsiToUnicode(LPCSTR pszA, LPOLESTR* ppszW)
{

    ULONG cCharacters;
    DWORD dwError;

    // If input is null then just return the same.
    if (NULL == pszA)
    {
        *ppszW = NULL;
        return NOERROR;
    }

    // Determine number of wide characters to be allocated for the
    // Unicode string.
    cCharacters =  strlen(pszA)+1;

    // Use of the OLE allocator is required if the resultant Unicode
    // string will be passed to another COM component and if that
    // component will free it. Otherwise you can use your own allocator.
    *ppszW = (LPOLESTR) CoTaskMemAlloc(cCharacters*2);
    if (NULL == *ppszW)
        return E_OUTOFMEMORY;

    // Covert to Unicode.
    if (0 == MultiByteToWideChar(CP_ACP, 0, pszA, cCharacters,
                  *ppszW, cCharacters))
    {
        dwError = GetLastError();
        CoTaskMemFree(*ppszW);
        *ppszW = NULL;
        return HRESULT_FROM_WIN32(dwError);
    }

    return NOERROR;
}

/*
 * UnicodeToAnsi converts the Unicode string pszW to an ANSI string
 * and returns the ANSI string through ppszA. Space for the
 * the converted string is allocated by UnicodeToAnsi.
 */ 

HRESULT __fastcall UnicodeToAnsi(LPCOLESTR pszW, LPSTR* ppszA)
{

    ULONG cbAnsi, cCharacters;
    DWORD dwError;

    // If input is null then just return the same.
    if (pszW == NULL)
    {
        *ppszA = NULL;
        return NOERROR;
    }

    cCharacters = wcslen(pszW)+1;
    // Determine number of bytes to be allocated for ANSI string. An
    // ANSI string can have at most 2 bytes per character (for Double
    // Byte Character Strings.)
    cbAnsi = cCharacters*2;

    // Use of the OLE allocator is not required because the resultant
    // ANSI  string will never be passed to another COM component. You
    // can use your own allocator.
    *ppszA = (LPSTR) CoTaskMemAlloc(cbAnsi);
    if (NULL == *ppszA)
        return E_OUTOFMEMORY;

    // Convert to ANSI.
    if (0 == WideCharToMultiByte(CP_ACP, 0, pszW, cCharacters, *ppszA,
                  cbAnsi, NULL, NULL))
    {
        dwError = GetLastError();
        CoTaskMemFree(*ppszA);
        *ppszA = NULL;
        return HRESULT_FROM_WIN32(dwError);
    }
    return NOERROR;

}

//
 // CStringA to CStringW
 //
 CStringW CStrA2CStrW(const CStringA &cstrSrcA)
 {
     int len = MultiByteToWideChar(CP_ACP, 0, LPCSTR(cstrSrcA), -1, NULL, 0);
     wchar_t *wstr = new wchar_t[len];
     memset(wstr, 0, len*sizeof(wchar_t));
     MultiByteToWideChar(CP_ACP, 0, LPCSTR(cstrSrcA), -1, wstr, len);
     CStringW cstrDestW = wstr;
     delete[] wstr;
 
     return cstrDestW;
 }
 
 //
 // CStringW to CStringA
 //
 CStringA CStrW2CStrA(const CStringW &cstrSrcW)
 {
     int len = WideCharToMultiByte(CP_ACP, 0, LPCWSTR(cstrSrcW), -1, NULL, 0, NULL, NULL);
     char *str = new char[len];
     memset(str, 0, len);
     WideCharToMultiByte(CP_ACP, 0, LPCWSTR(cstrSrcW), -1, str, len, NULL, NULL);
     CStringA cstrDestA = str;
     delete[] str;
 
     return cstrDestA;
 }

 CString DateStr(CString &date)
 {
	 CString str;
	 int findpos = 0, copypos = 0;
	 wchar_t c;
	 if(date.Find('/') != -1){
		 c = '/';
	 }else if(date.Find('-') != -1){
		 c = '-';
	 }else{
		 return str;
	 }
	 int len = date.GetLength();

	 findpos = date.Find(c);
	 for(int i = copypos; i < findpos; i++){
		 str += date[i];
	 }
	 str += L"Äê";
	 
	 copypos = findpos + 1;
	 findpos = date.Find(c, copypos);
	 for(int i = copypos; i < findpos; i++){
		 str += date[i];
	 }
	 str += L"ÔÂ";

	 copypos = findpos + 1;
	 //findpos = date.Find(c, copypos);
	 for(int i = copypos; i < len; i++){
		 str += date[i];
	 }
	 str += L"ÈÕ";
	 return str;
 }

 int CutZeros(CString &str)
 { 
	 if(str.Find('.') == -1){
		return 0;
	 }

	 CString tmp;
	 int size = str.GetLength();
	 int flag = 0;
	 for(int i = size - 1; i >= 0; i--){
		 if(flag == 0){
			 if(str[i] == '0'){
				 continue;
			 }else  if(str[i] == '.'){
				 flag = 1;
				 continue;
			 }else{
				 flag = 1;
			 }
		 }
		 tmp = str[i] + tmp;		 
	 }
	 str = tmp;
	 return 0;
 }