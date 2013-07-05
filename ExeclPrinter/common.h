#ifndef _COMMOM_H_
#define _COMMOM_H_

#include "stdafx.h"
//
 // CStringA to CStringW
 //
 CStringW CStrA2CStrW(const CStringA &cstrSrcA);

  //
 // CStringW to CStringA
 //
 CStringA CStrW2CStrA(const CStringW &cstrSrcW);

HRESULT __fastcall AnsiToUnicode(LPCSTR pszA, LPOLESTR* ppszW);
HRESULT __fastcall UnicodeToAnsi(LPCOLESTR pszW, LPSTR* ppszA);

CString DateStr(CString &date);
int CutZeros(CString &str);

#endif