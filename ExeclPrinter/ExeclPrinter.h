
// ExeclPrinter.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CExeclPrinterApp:
// �йش����ʵ�֣������ ExeclPrinter.cpp
//

class CExeclPrinterApp : public CWinApp
{
public:
	CExeclPrinterApp();

// ��д
public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CExeclPrinterApp theApp;