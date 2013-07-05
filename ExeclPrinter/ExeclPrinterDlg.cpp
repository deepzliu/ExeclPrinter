


// ExeclPrinterDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "ExeclPrinter.h"
#include "ExeclPrinterDlg.h"
#include "afxdialogex.h"
//#include "excel9.h"
#include "common.h"
#include <vector>
#include <string>
#include <fstream>
using namespace std;
using namespace std;

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CExeclPrinterDlg 对话框

CExeclPrinterDlg::CExeclPrinterDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CExeclPrinterDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

CExeclPrinterDlg::~CExeclPrinterDlg()
{
	m_ColumnsInfo.clear();
	m_TableMap.clear();	

	app_t.SetAlertBeforeOverwriting(false);
	app_t.SetDisplayAlerts(false);
	book_t.Close(_variant_t(FALSE),_variant_t(m_TempletPath),_variant_t(FALSE));
	books_t.Close();
	app_t.Quit();
	range_t.ReleaseDispatch(); 
	sheet_t.ReleaseDispatch(); 
	sheets_t.ReleaseDispatch(); 
	book_t.ReleaseDispatch(); 	
	books_t.ReleaseDispatch();	
	//app_t.SetVisible(true);
	app_t.ReleaseDispatch();

	app_s.SetAlertBeforeOverwriting(false);
	app_s.SetDisplayAlerts(false);
	book_s.Close(_variant_t(FALSE),_variant_t(m_SourcePath),_variant_t(FALSE));//先关闭退出，再release，顺序不能反。
	books_s.Close();
	app_s.Quit();	
	range_s.ReleaseDispatch();
	sheet_s.ReleaseDispatch(); 
	sheets_s.ReleaseDispatch();	
	book_s.ReleaseDispatch(); 
	books_s.ReleaseDispatch();
	//app_s.SetVisible(true);
	app_s.ReleaseDispatch();	

	CoUninitialize();
}

void CExeclPrinterDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, m_ContentList);
	//DDX_Control(pDX, IDC_LIST2, m_SubList);
	DDX_Control(pDX, IDC_EDIT1, m_SearchText);
}

BEGIN_MESSAGE_MAP(CExeclPrinterDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//ON_BN_CLICKED(IDC_BUTTON_EXCEL, &CExeclPrinterDlg::OnBnClickedButtonExcel)
	ON_BN_CLICKED(IDC_BUTTON1, &CExeclPrinterDlg::OnSearch)
	ON_BN_CLICKED(IDC_BUTTON4, &CExeclPrinterDlg::OnFirstPage)
	ON_BN_CLICKED(IDC_BUTTON2, &CExeclPrinterDlg::OnPrePage)
	ON_BN_CLICKED(IDC_BUTTON3, &CExeclPrinterDlg::OnNextPage)
	ON_BN_CLICKED(IDC_BUTTON5, &CExeclPrinterDlg::OnRefresh)
	ON_BN_CLICKED(IDC_PRINT, &CExeclPrinterDlg::OnBnClickedPrint)
END_MESSAGE_MAP()


// CExeclPrinterDlg 消息处理程序

BOOL CExeclPrinterDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
	if (CoInitialize(NULL)!=0) 
	{ 
		AfxMessageBox(L"初始化COM支持库失败!"); 
		exit(1); 
	}

	MyInit();

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CExeclPrinterDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CExeclPrinterDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CExeclPrinterDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CExeclPrinterDlg::OnBnClickedButtonExcel()
{
	// TODO: 在此添加控件通知处理程序代码
	SetExcel();
}


int CExeclPrinterDlg::SetExcel()
{
#if 0
	_Application app;    
	Workbooks books;
	_Workbook book;
	Worksheets sheets;
	_Worksheet sheet;
	Range range;

	CString templet = L"templet1.xlt";

	//创建Excel 2000服务器(启动Excel) 
	if (!app.CreateDispatch(L"Excel.Application",NULL)) 
	{ 
		AfxMessageBox(L"创建Excel服务失败!"); 
		exit(1); 
	} 
	app.SetVisible(false); 
	//利用模板文件建立新文档 
	wchar_t path[MAX_PATH];
	GetCurrentDirectory(MAX_PATH,path);
	CString strPath = path;
	strPath += L"\\" + templet;
	books.AttachDispatch(app.GetWorkbooks(),true);
	book.AttachDispatch(books.Add(_variant_t(strPath)));
	//得到Worksheets 
	sheets.AttachDispatch(book.GetWorksheets(),true);	
	//得到sheet1 
	sheet.AttachDispatch(sheets.GetItem(_variant_t("sheet1")),true);
	CString str1;
	str1 = L"第1页";
	sheet.SetName(str1);
	for( int i=0;i<sheets.GetCount()-1;i++)
	{
		sheet = sheet.GetNext();
		str1.Format(L"第%d页",i+2);
		sheet.SetName(str1);
	}
	sheet.AttachDispatch(sheets.GetItem(_variant_t("第1页")),true);
	//得到全部Cells，此时,rgMyRge是cells的集合 
	range.AttachDispatch(sheet.GetCells(),true);

	
	
	int row_start = 21, colume_start = 3;
	int row = row_start + 4, colume = colume_start + 1;
	//wchar_t list[4][64] = {L"客户名称", L"客户 ID", L"帐单号", L"日期"};
	//wchar_t list2[4][64] = {"客户名称", "客户 ID", "帐单号", "日期"};
	wchar_t list[4][64] = {L"清华同方", L"88508", L"10101345", L"=C9"};
	char list5[4][64] = {"清华同方", "88508", "10101345", "=C9"};
	//char *tmp[64];
	CStringA strA;
	for(int i = row_start; i < row; i++){
		for(int j = colume_start; j < colume; j++){
			//UnicodeToAnsi(list[i], tmp);
			//strA = list5[i - row_start];
			range.SetItem(_variant_t(i), _variant_t(j), _variant_t(list5[i - row_start]));
		}
	}
	colume_start = 2;
	_variant_t vt;
	CString str;
	vector<CString> list3;
	for(int i = row_start; i < row; i++){
		for(int j = colume_start; j < colume; j++){
			vt = range.GetItem(_variant_t(i), _variant_t(j));
			//wcscpy(list2[i],vt.bstrVal);
			str = vt.bstrVal;
			vt.Clear();
			list3.push_back(str);
		}
	}
	app.SetVisible(true);
	//book.PrintPreview(_variant_t(false));
	
	//释放对象 
	range.ReleaseDispatch(); 
	sheet.ReleaseDispatch(); 
	sheets.ReleaseDispatch(); 
	book.ReleaseDispatch(); 
	books.ReleaseDispatch();
	app.ReleaseDispatch(); 
#endif
	return 0;
}

void CExeclPrinterDlg::ParseListShow(string &listshow)
{
	int size = listshow.size();
	int pos = 0;
	Column_Info ci;
	string sub, line;
	line = listshow;
	while((pos = line.find(',')) >= 0){
		if(pos == size - 1) break;
		ci.column = 0;
		ci.width = 0;
		sub = line.substr(0, pos - 0);
		line = line.substr(pos + 1, line.size() - pos);
		ci.column = atoi(sub.c_str());
		pos = sub.find(':');
		if(pos > -1){
			sub = sub.substr(pos+1);
			ci.width = atoi(sub.c_str());
		}
		m_ColumnsInfo.push_back(ci);
	}
}

void CExeclPrinterDlg::ParseMap(string &map)
{
	int size = map.size();
	int pos = 0, v[9], n = 0;
	Table_Map ci;
	string sub, line;
	line = map;
	while((pos = line.find(',')) >= 0){
		if(pos == size - 1) break;
		memset(&ci, 0, sizeof(ci));
		sub = line.substr(0, pos - 0);
		line = line.substr(pos + 1, line.size() - pos);
		v[n++] = atoi(sub.c_str());
		pos = sub.find('-');
		if(pos > -1){
			sub = sub.substr(pos+1);
			v[n++] = atoi(sub.c_str());
			if(v[n-1] == 0) v[n-1] = 0x7fffffff;
		}else{
			if(n < 8){
				v[n] = v[n-1];
				n++;
			}
		}
	}

	ci.dx0 = v[0];
	ci.dx1 = v[1];
	ci.dy0 = v[2];
	ci.dy1 = v[3];
	ci.sx0 = v[4];
	ci.sx1 = v[5];
	ci.sy0 = v[6];
	ci.sy1 = v[7];
	ci.flag = v[8];
	m_TableMap.push_back(ci);
}

void CExeclPrinterDlg::ReadConfig()
{
	string config = "config.ini";
	string map = "map.ini";
	int pos = 0;
	string head, tail;

	//config.ini
	m_Tempetsheet = "Sheet1";
	m_Sourcesheet = "Sheet1";
	m_PageInfo.count_per_page = 24;
	//m_Rowsperpage = 20;
	ifstream fin(config);
	if(fin.is_open()){
		string line;
		while(getline(fin, line)){
			pos = line.find('=');
			head = line.substr(0, pos);
			tail = line.substr(pos+1, line.size() - pos - 1);
			if(head == "title"){
				m_Title.Format("%s",tail.c_str());
			}else if(head == "templet"){
				m_TempletExcel = tail;
			}else if(head == "templetsheet"){
				m_Tempetsheet = tail;
			}else if(head == "datasource"){
				m_DataSource = tail;
			}else if(head == "sourcesheet"){
				m_Sourcesheet = tail;
			}else if(head == "listshow"){
				ParseListShow(tail);
			}else if(head == "rowsperpage"){
				m_PageInfo.count_per_page = atoi(tail.c_str());
			}else if(head == "map"){
				ParseMap(tail);
			}

			
		}
		fin.close();
	}

}

void CExeclPrinterDlg::InsertListColumns()
{
	//RECT rect;
	//m_ContentList.GetWindowRect(&rect);
	//
	////int width = rect.bottom - rect.top;
	//int height = rect.right - rect.left;
	
	CString str;
	VARIANT vt;
	Range rg;	

	//long countflag = range_s.GetCount();
	int size = m_ColumnsInfo.size();
	int add = 0;
	m_ContentList.SetExtendedStyle(LVS_EX_FLATSB
		|LVS_EX_FULLROWSELECT
		|LVS_EX_HEADERDRAGDROP
		//|LVS_EX_ONECLICKACTIVATE
		|LVS_EX_GRIDLINES);
	m_ContentList.InsertColumn(0, L"N.", LVCFMT_LEFT, 25);
	for(int i = 0; i < size; i++){
		//if(countflag > 0){
		rg.AttachDispatch(range_s.GetItem(_variant_t((1)), _variant_t((m_ColumnsInfo[i].column + add))).pdispVal, true);	
		vt = rg.GetValue();
		str = vt.bstrVal;
		rg.ReleaseDispatch();
		if(str == L"Year" || str == L"Month" || str == L"Date"){
			i--;
			add++;
		}else{
			m_ContentList.InsertColumn(i+1, str, LVCFMT_LEFT, m_ColumnsInfo[i].width);		
		}
	}


}

void CExeclPrinterDlg::CountRows(Range &range)
{
	Range rg;
	VARIANT vt;
	CString str;
	int n = 2, flag = 0;
	while(flag <= 10){
		rg.AttachDispatch(range.GetItem(_variant_t(n), _variant_t(1)).pdispVal, true);	
		vt = rg.GetValue();
		rg.ReleaseDispatch();
		if(vt.vt != 0 ){
			flag = 0;
			m_PageInfo.count_Items++;
		}else{
			flag++;
		}
		n++;
	}
}

void CExeclPrinterDlg::OpenTemplet()
{
	//Range rg;

	//创建Excel服务器(启动Excel) 
	if (!app_t.CreateDispatch(L"Excel.Application",NULL)) 
	{ 
		int err = GetLastError();
		
		AfxMessageBox(L"创建Excel服务失败!"); 
		exit(1); 
	} 
	app_t.SetVisible(false); 
	//利用模板文件建立新文档 
	char path[MAX_PATH];
	GetCurrentDirectoryA(MAX_PATH,path);
	string strPath = path;
	strPath += "\\";
	strPath += m_TempletExcel;
	m_TempletPath.Format("%s",strPath.c_str()); 
	books_t.AttachDispatch(app_t.GetWorkbooks(),true);
	book_t.AttachDispatch(books_t.Add(_variant_t(strPath.c_str())));
	//得到Worksheets 
	sheets_t.AttachDispatch(book_t.GetWorksheets(),true);	
	//得到sheet1 
	sheet_t.AttachDispatch(sheets_t.GetItem(_variant_t(m_Tempetsheet.c_str())),true);
	//得到全部Cells，此时,rgMyRge是cells的集合 
	range_t.AttachDispatch(sheet_t.GetCells(),true);
	//err = GetLastError();
}

void CExeclPrinterDlg::MyInit()
{
	int err = 0;
	m_RowHeight = 25;
	memset(&m_PageInfo, 0, sizeof(m_PageInfo));

	ReadConfig();
	m_RowdefaultHeight = 20;
	SetWindowTextA(this->m_hWnd, m_Title);

	// for templet
	OpenTemplet();

	// for data source 
	//创建Excel服务器(启动Excel) 
	if (!app_s.CreateDispatch(L"Excel.Application",NULL)) 
	{ 
		AfxMessageBox(L"创建Excel服务失败!"); 
		exit(1); 
	} 
	app_s.SetVisible(false); 
	//利用模板文件建立新文档 
	char path[MAX_PATH];
	GetCurrentDirectoryA(MAX_PATH,path);
	string strPath = path;
	strPath += "\\";
	strPath += m_DataSource;
	m_SourcePath.Format("%s",strPath.c_str()); 
	books_s.AttachDispatch(app_s.GetWorkbooks(),true);
	book_s.AttachDispatch(books_s.Add(_variant_t(strPath.c_str())));
	//得到Worksheets 
	sheets_s.AttachDispatch(book_s.GetWorksheets(),true);	
	//得到sheet1 
	sheet_s.AttachDispatch(sheets_s.GetItem(_variant_t(m_Sourcesheet.c_str())),true);
	//得到全部Cells，此时,rgMyRge是cells的集合 
	range_s.AttachDispatch(sheet_s.GetCells(),true);
	//err = GetLastError();

	InsertListColumns();

	CountRows(range_s);

	for(int i = 0; i < m_PageInfo.count_per_page; i++){
		m_ContentList.InsertItem(i,0);
		//m_ContentList.SetItemText(i, 0, L"abc");
	}
	DisplayPage();

}

int CExeclPrinterDlg::GetDataType(int column)
{
	//int flag = 0;
	int size = m_TableMap.size();
	for(int i = 0; i < size; i++){
		if(column == m_TableMap[i].sx0){
			return m_TableMap[i].flag;
		}
	}
	return 0;
}

void CExeclPrinterDlg::DisplayPage()
{
	DisplayPage(m_PageInfo.curr_page, m_PageInfo.count_per_page, m_PageInfo.count_Items, 0);
	//m_ContentList.c
}

/*
	nth:	第几页
	nItem:	窗口表格每页行数
	nCount：总共行数
	searchflag: 0,不搜索；1，搜索；
*/
void CExeclPrinterDlg::DisplayPage(int nth, int nItem, int nCount, int searchflag)
{
	if(nth > (nCount + nItem - 1)/nItem) return;//显示的行数大于总行数的情况

	CString search;
	CStringArray strarr;
	CString str;//('\0', 1024);
	int datatype = 0;
	CTime t;
	VARIANT vt;
	Range rg;
	int add = 0, nflag = 0, flag = 0, matchedflag = 0;
	int row = nth * nItem + 2;//包括一行标题需要去掉
	int column = m_ColumnsInfo.size();
	strarr.SetSize(column);
	m_PageInfo.count_real = 0;
	if(searchflag == 1){
		m_SearchText.GetWindowText(search);
		if(search.IsEmpty()) searchflag = 0;
	}

	for(int i = 0; i < nItem; i++){
		
		str.Format(L"%d", i+1);
		m_ContentList.SetItemText(i, 0, str);
		str.ReleaseBuffer();
		flag = 0;
		matchedflag = 0;
		for(int j = 0; j < column; j++){
			rg.AttachDispatch(range_s.GetItem(_variant_t(row+i), _variant_t(m_ColumnsInfo[j].column + add)).pdispVal, true);	
			vt = rg.GetValue();
			
			switch(vt.vt){
			case VT_DATE:
				/*add += 3;
				rg.AttachDispatch(range_s.GetItem(_variant_t(row+i), _variant_t(m_ColumnsInfo[j].column + 1)).pdispVal, true);	
				vt = rg.GetValue();
				str.Format(L"%ld", (long)vt.dblVal);
				rg.AttachDispatch(range_s.GetItem(_variant_t(row+i), _variant_t(m_ColumnsInfo[j].column + 2)).pdispVal, true);	
				vt = rg.GetValue();
				str.Format(L"%s/%02ld", str.GetBuffer(), (long)vt.dblVal);
				rg.AttachDispatch(range_s.GetItem(_variant_t(row+i), _variant_t(m_ColumnsInfo[j].column + 3)).pdispVal, true);	
				vt = rg.GetValue();*/
				
				//Excel的日期是以1900-1-1为基准，CTime是以1970-1-1为基准
				//t = CTime((long long)vt.dblVal * 24 * 60 * 60);
				//str.Format(L"%d年%02d月%02d日", t.GetYear() - 70, t.GetMonth(), t.GetDay() - 1);
				
				vt = rg.GetText();
				str = vt.bstrVal;
				//str = "2012-2-2";
				str = DateStr(str);
				break;//7
			case VT_BSTR:
				str = vt.bstrVal;
				break;//8
			case VT_I2://2
				str.Format(L"%d", vt.iVal);
				break;
			case VT_R8://5
				datatype = GetDataType(j+1);
				if(datatype == 1){
					str.Format(L"%llf", vt.dblVal);
					CutZeros(str);
					//str.Format(L"%g", vt.dblVal);
				}else{
					str.Format(L"%lld", (long long)vt.dblVal);
				}
				break;
			case VT_I4://3
				str.Format(L"%ld", vt.lVal);
				break;
			case VT_EMPTY:
				str = "";
				if(j == 0){
					flag = 1;
					nflag++;
					if(nflag >= 5){
						rg.ReleaseDispatch();
						goto exit0;
					}
				}
				break;
			default:
				str = "";
				break;
			}
			strarr[j] = str;
			if(searchflag == 1){				
				if(matchedflag == 0 && str.Find(search) != -1){
					matchedflag = 1;
				}
			}
			rg.ReleaseDispatch();			
			//str.ReleaseBuffer();
		}
		
		if(searchflag == 1 && matchedflag == 1){
			m_PageInfo.count_real++;
			for(int j = 0; j < column; j++){
				m_ContentList.SetItemText(i, j+1, strarr[j]);
			}
		}else if(searchflag == 0){
			for(int j = 0; j < column; j++){
				m_ContentList.SetItemText(i, j+1, strarr[j]);
			}
			if(flag == 0) m_PageInfo.count_real++;
		}else if(searchflag == 1 && matchedflag == 0){
			i--;
			row++;
		}
		add = 0;
	}

exit0:
	str = "";
	nflag = m_PageInfo.count_per_page - m_PageInfo.count_real;
	for(int i = 0; i < nflag; i++){
		for(int j = 0; j < column; j++){
			m_ContentList.SetItemText(m_PageInfo.count_real + i, j+1, str);
		}
	}

	strarr.RemoveAll();
}

void CExeclPrinterDlg::OnSearch()
{
	// TODO: 在此添加控件通知处理程序代码
	m_PageInfo.curr_page = 0;

	CString str = L"";
	int column = m_ColumnsInfo.size();
	int nflag = m_PageInfo.count_per_page;
	for(int i = 0; i < nflag; i++){
		for(int j = 0; j < column; j++){
			m_ContentList.SetItemText(i, j+1, str);
		}
	}

	DisplayPage(m_PageInfo.curr_page, m_PageInfo.count_per_page, m_PageInfo.count_Items, 1);

}


void CExeclPrinterDlg::OnFirstPage()
{
	// TODO: 在此添加控件通知处理程序代码
	m_PageInfo.curr_page = 0;
	DisplayPage();
}


void CExeclPrinterDlg::OnPrePage()
{
	// TODO: 在此添加控件通知处理程序代码
	if(m_PageInfo.curr_page <= 0){
		MessageBox(L"已经是第一页", L"提示", MB_ICONWARNING);
	}else{
		m_PageInfo.curr_page--;
		DisplayPage();
	}
}


void CExeclPrinterDlg::OnNextPage()
{
	// TODO: 在此添加控件通知处理程序代码
	//if(m_PageInfo.curr_page * m_PageInfo.count_per_page >= m_PageInfo.count_Items - m_PageInfo.count_per_page + 1){
	if(m_PageInfo.count_real < m_PageInfo.count_per_page){
		MessageBox(L"已经是最后一页", L"提示", MB_ICONWARNING);
	}else{
		m_PageInfo.curr_page++;
		DisplayPage();
	}
}


void CExeclPrinterDlg::OnRefresh()
{
	// TODO: 在此添加控件通知处理程序代码
}


void CExeclPrinterDlg::OnBnClickedPrint()
{
	// TODO: 在此添加控件通知处理程序代码
	CString str;
	CStringArray strarr;

	//需要决断模板是否打开
	//OpenTemplet();

	TCHAR szBuf[1024];
    LVITEM lvi;
    //lvi.iItem = 0;
    //lvi.iSubItem = 0;
    lvi.mask = LVIF_TEXT;
    lvi.pszText = szBuf;
    lvi.cchTextMax = 1024;
	int cSize = m_ColumnsInfo.size();
	strarr.SetSize(cSize);
	int flag = 0;
	// get selected row data for printing
	for(int i = 0; i < m_PageInfo.count_real; i++){
		if( m_ContentList.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED ){
			flag = 1;
			lvi.iItem = i;			
			for(int j = 1; j < cSize; j++){
				lvi.iSubItem = j;
				m_ContentList.GetItem(&lvi);
				strarr[j-1] = lvi.pszText;
			}
			break;
		}
	}
	if(flag == 0){
		MessageBox(L"请选择一行", L"提示", MB_ICONWARNING);
		strarr.RemoveAll();
		return ;
	}

	//暂只考虑一对一的map情况，不考虑批量对应
	// set the data above getted to the templet
	int tSize = m_TableMap.size();
	Range rg;
	VARIANT vt;
	double width = 0, height = 0;
	for(int i = 0; i < tSize; i++){
		if(m_TableMap[i].sx0 < cSize){
			range_t.SetItem(_variant_t(m_TableMap[i].dx0), _variant_t(m_TableMap[i].dy0), _variant_t(strarr[m_TableMap[i].sx0 - 1].GetBuffer()));
			
			rg.AttachDispatch(range_t.GetItem(_variant_t(m_TableMap[i].dx0), _variant_t(m_TableMap[i].dy0)).pdispVal, true);
			vt = rg.GetHeight();
			height = vt.dblVal;
			if(height <= m_RowdefaultHeight && strarr[m_TableMap[i].sx0 - 1].GetLength() > 12){
				rg.SetRowHeight(_variant_t(height * 2));
			}
			rg.ReleaseDispatch();
		}
	}	

	//print
	app_t.SetVisible(true);
	book_t.PrintPreview(_variant_t(false));
	app_t.SetVisible(false);
	strarr.RemoveAll();
}
