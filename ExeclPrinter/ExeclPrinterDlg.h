
// ExeclPrinterDlg.h : ͷ�ļ�
//

#pragma once
#include "afxcmn.h"
#include <vector>
#include "excel9.h"
#include "afxwin.h"
using namespace std;

typedef struct _Table_Map{
	int dx0, dy0; //Ŀ�������ʼλ�ã�һ��Ŀ����excel
	int dx1, dy1; //Ŀ����Ľ���λ�ã������������Ӧ��������ʼλ����ͬ
	int sx0, sy0; //Դ������ʼλ�ã�Ŀǰֻ����excel�����������������ݿ�ķ������
	int sx1, sy1; //Դ���Ľ���λ�ã������������Ӧ��������ʼλ����ͬ
	int flag;//�������ͱ�ǣ�0Ϊ������1Ϊ��������������С�������λ����2Ϊ�ı���3Ϊʱ�䣻
}Table_Map;

typedef struct _Column_Info{//�����ڴ����б���˳����ʾ���м�����
	int column;
	int width;
}Column_Info;

typedef struct _Page_Info{
	long count_Items;
	long curr_page;
	long count_per_page;
	long count_real;
}Page_Info;

// CExeclPrinterDlg �Ի���
class CExeclPrinterDlg : public CDialogEx
{
// ����
public:
	CExeclPrinterDlg(CWnd* pParent = NULL);	// ��׼���캯��
	~CExeclPrinterDlg();

// �Ի�������
	enum { IDD = IDD_EXECLPRINTER_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:

	int SetExcel();

	afx_msg void OnBnClickedButtonExcel();
	void ReadConfig();
	void MyInit();
	void InsertListColumns();
	CStringA m_Title;
	string m_TempletExcel;
	CStringA m_TempletPath;
	string m_Tempetsheet;
	string m_DataSource;
	CStringA m_SourcePath;
	string m_Sourcesheet;
	CListCtrl m_ContentList;
	//CListCtrl m_SubList;

	//list table's size
	double m_RowdefaultHeight;
	int m_RowHeight;
	int m_Rowsperpage;
	vector<Column_Info> m_ColumnsInfo;
	void ParseListShow(string &listshow);	

	//list table's map
	vector<Table_Map> m_TableMap;
	void ParseMap(string &map);
	int GetDataType(int column);

	// for templet excel
	_Application app_t;    
	Workbooks books_t;
	_Workbook book_t;
	Worksheets sheets_t;
	_Worksheet sheet_t;
	Range range_t;

	// for data source excel
	_Application app_s;    
	Workbooks books_s;
	_Workbook book_s;
	Worksheets sheets_s;
	_Worksheet sheet_s;
	Range range_s;
	void OpenTemplet();

	// list page info
	Page_Info m_PageInfo;
	void CountRows(Range &range);
	void DisplayPage();
	void DisplayPage(int nth, int nItem, int nCount, int searchflag);
	CEdit m_SearchText;

	void AppRelease();
	afx_msg void OnSearch();
	afx_msg void OnFirstPage();
	afx_msg void OnPrePage();
	afx_msg void OnNextPage();
	afx_msg void OnRefresh();
	afx_msg void OnBnClickedPrint();
};
