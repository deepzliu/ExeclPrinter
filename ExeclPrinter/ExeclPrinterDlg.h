
// ExeclPrinterDlg.h : 头文件
//

#pragma once
#include "afxcmn.h"
#include <vector>
#include "excel9.h"
#include "afxwin.h"
using namespace std;

typedef struct _Table_Map{
	int dx0, dy0; //目标表格的起始位置，一般目标是excel
	int dx1, dy1; //目标表格的结束位置，如果非批量对应，则与起始位置相同
	int sx0, sy0; //源表格的起始位置，目前只考虑excel，尽可能往兼容数据库的方向设计
	int sx1, sy1; //源表格的结束位置，如果非批量对应，则与起始位置相同
	int flag;//数据类型标记，0为整数，1为浮点数（均保留小数点后两位），2为文本，3为时间；
}Table_Map;

typedef struct _Column_Info{//描述在窗体列表中顺序显示的列及其宽度
	int column;
	int width;
}Column_Info;

typedef struct _Page_Info{
	long count_Items;
	long curr_page;
	long count_per_page;
	long count_real;
}Page_Info;

// CExeclPrinterDlg 对话框
class CExeclPrinterDlg : public CDialogEx
{
// 构造
public:
	CExeclPrinterDlg(CWnd* pParent = NULL);	// 标准构造函数
	~CExeclPrinterDlg();

// 对话框数据
	enum { IDD = IDD_EXECLPRINTER_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
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
