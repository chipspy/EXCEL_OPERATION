
// mfc069_excelDlg.cpp: 구현 파일
//

#include "pch.h"
#include "framework.h"
#include "mfc069_excel.h"
#include "mfc069_excelDlg.h"
#include "afxdialogex.h"

#include "XLEzAutomation.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// Cmfc069excelDlg 대화 상자



Cmfc069excelDlg::Cmfc069excelDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_MFC069_EXCEL_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void Cmfc069excelDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, m_ListCtrl1);
}

BEGIN_MESSAGE_MAP(Cmfc069excelDlg, CDialogEx)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_EXCEL_READ_BTN, &Cmfc069excelDlg::OnBnClickedExcelReadBtn)
	ON_BN_CLICKED(IDC_EXCEL_WRITE_BTN, &Cmfc069excelDlg::OnBnClickedExcelWriteBtn)
END_MESSAGE_MAP()


// Cmfc069excelDlg 메시지 처리기

BOOL Cmfc069excelDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 이 대화 상자의 아이콘을 설정합니다.  응용 프로그램의 주 창이 대화 상자가 아닐 경우에는
	//  프레임워크가 이 작업을 자동으로 수행합니다.
	SetIcon(m_hIcon, TRUE);			// 큰 아이콘을 설정합니다.
	SetIcon(m_hIcon, FALSE);		// 작은 아이콘을 설정합니다.

	//1. 이미지 리스트 생성과 초기화
	CImageList	ilLarge, ilSmall;
	ilLarge.Create(32, 32, ILC_COLOR4, 1, 1);
	ilSmall.Create(16, 16, ILC_COLOR4, 1, 1);
	//2. 이미지 리스트 설정
	m_ListCtrl1.SetImageList(&ilLarge, LVSIL_NORMAL);
	m_ListCtrl1.SetImageList(&ilSmall, LVSIL_SMALL);
	//3. 분리
	ilLarge.Detach();
	ilSmall.Detach();

	//1. 리스트 컨트롤 배경색 설정
	COLORREF crBkColor = (RGB(235, 254, 211));
	m_ListCtrl1.SetTextBkColor(crBkColor);
	//ASSERT(m_MenuList.GetTextBkColor() == crBkColor); // assert C
	//m_ListCtrl1.GetTextBkColor() == crBkColor;
	// 중요: 프로젝트 속성=> 문자집합=> 멀티바이트로 설정해야 리스트 컨트롤 칼
	m_ListCtrl1.InsertColumn(0, _T("순"), LVCFMT_LEFT, 150);
	m_ListCtrl1.InsertColumn(1, _T("정식과목명"), LVCFMT_CENTER, 100);
	m_ListCtrl1.InsertColumn(2, _T("단축과목명"), LVCFMT_CENTER, 100);
	m_ListCtrl1.InsertColumn(3, _T("교사명"), LVCFMT_CENTER, 100);
	m_ListCtrl1.InsertColumn(4, _T("1-1반"), LVCFMT_CENTER, 60);
	m_ListCtrl1.InsertColumn(5, _T("1-2반"), LVCFMT_CENTER, 60);
	m_ListCtrl1.InsertColumn(6, _T("2-1반"), LVCFMT_CENTER, 60);
	m_ListCtrl1.InsertColumn(7, _T("2-2반"), LVCFMT_CENTER, 60);
	m_ListCtrl1.InsertColumn(8, _T("3-1반"), LVCFMT_CENTER, 60);
	m_ListCtrl1.InsertColumn(9, _T("3-2반"), LVCFMT_CENTER, 60);


	// 확장 스타일 지정
	m_ListCtrl1.SetExtendedStyle(m_ListCtrl1.GetExtendedStyle() | LVS_EX_GRIDLINES | LVS_EX_TRACKSELECT | LVS_EX_CHECKBOXES);
		// SetExtendedStyle: 확장스타일,
		// LVS_EX_GRIDLINES: List View ExtendedStyle 격자표시
		// LVS_EX_CHECKBOXES: 체크박스표시
		// LVS_EX_TRACKSELECT: 마우스로 아이템을 클릭하지 않아도 마우스를 갖다 대면 자동으로 선택
	m_ListCtrl1.SetExtendedStyle(m_ListCtrl1.GetExtendedStyle() | LVS_EX_FULLROWSELECT | LVS_EX_HEADERDRAGDROP);
	    //LVS_EX_FULLROWSELECT); // o
	     // LVS_EX_HEADERDRAGDROP://2
	return TRUE;  // 포커스를 컨트롤에 설정하지 않으면 TRUE를 반환합니다.
}

// 대화 상자에 최소화 단추를 추가할 경우 아이콘을 그리려면
//  아래 코드가 필요합니다.  문서/뷰 모델을 사용하는 MFC 애플리케이션의 경우에는
//  프레임워크에서 이 작업을 자동으로 수행합니다.

void Cmfc069excelDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 그리기를 위한 디바이스 컨텍스트입니다.

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 클라이언트 사각형에서 아이콘을 가운데에 맞춥니다.
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 아이콘을 그립니다.
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// 사용자가 최소화된 창을 끄는 동안에 커서가 표시되도록 시스템에서
//  이 함수를 호출합니다.
HCURSOR Cmfc069excelDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void Cmfc069excelDlg::OnBnClickedExcelReadBtn()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	CString m_strFileName;
	CString strData;
	CString data1, data2, data3, data4, data5;
	CString data6, data7, data8, data9, data10;
	CString str_data1, str_data5, str_data6, str_data7, str_data8, str_data9, str_data10; //문자 -> 숫자
	int int_data1, int_data5, int_data6, int_data7, int_data8, int_data9, int_data10; //  숫자 -> 문자
	

	CXLEzAutomation XL(FALSE); // 엑셀 함수를 사용하기 위한 클래스 변수 선언
	// TRUE: 열기
	CFileDialog dlg(TRUE, _T("*.xls"), NULL, OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST,
		_T("Excel Files (*.xls)|*.xls||"));
	//CFileDialog dlg(TRUE, L"*.*", NULL, OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST, 
	//	           _T("Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx|All Files (*.*)|*.*||"), NULL);
	//CFileDialog dlg(TRUE, _T("xlsx"), m_strFileName + _T(".xlsx"), OFN_HIDEREADONLY | OFN_NOCHANGEDIR, 
	//	                 // _T("xlsx (*.xlsx) | *.xlsx)"), NULL);
	//	_T("Excel 파일 (*.xlsx)|*.xlsx|모든 파일 (*.*)|*.*||"), NULL);
	// OFN_HIDEREADONLY: 읽기 전용 확인란을 숨깁니다.
    // OFN_NOCHANGEDIR: 사용자가 파일을 검색하는 동안 디렉터리를 변경한 경우 현재 디렉터리를 원래 값으로
	if (dlg.DoModal() == IDOK)
	{
		XL.OpenExcelFile(dlg.GetPathName()); // 엑셀 파일 열기
		m_ListCtrl1.DeleteAllItems(); // 리스트 컨트롤 화면 지우기
		for (int i = 2; i <= 11; i++) // 거꾸로 출력됨
		//for (int i = 1; i <= 2; i++) // 엑셀파일 데이터를 거꾸로 가져와야 vc++ 리스트컨트롤에 순서대로 출력
		{
			data1 = XL.GetCellValue(1, i); // 엑셀파일의 열, 행 값 가져오기
			data2 = XL.GetCellValue(2, i);
			data3 = XL.GetCellValue(3, i);
			data4 = XL.GetCellValue(4, i);
			data5 = XL.GetCellValue(5, i);
			data6 = XL.GetCellValue(6, i);
			data7 = XL.GetCellValue(7, i);
			data8 = XL.GetCellValue(8, i);
			data9 = XL.GetCellValue(9, i);
			data10 = XL.GetCellValue(10, i);

			// 리스트 컨트롤에 출력
			int_data1 = _ttoi(data1); // CString => int 9984|, str data1.Format(_T("%d"). int data1);
			str_data1.Format(L"%d", int_data1);
			m_ListCtrl1.InsertItem(i-2, str_data1); //순 
			m_ListCtrl1.SetItemText(i-2, 1, data2);  // 정식과목명
			m_ListCtrl1.SetItemText(i-2, 2, data3);  // 단축과목명
			m_ListCtrl1.SetItemText(i-2, 3, data4);  // 성명

			// _ttoi=_wttoi (유니코드 변환시)
			int_data5 = _ttoi(data5);   // 엑셀 숫자 자릿점 지우기 (CString -> int)
			str_data5.Format(L"%d", int_data5);
			int_data6 = _ttoi(data6);   // 엑셀 숫자 자릿점 지우기 (CString -> int)
			str_data6.Format(L"%d", int_data6);
			int_data7 = _ttoi(data7);   // 엑셀 숫자 자릿점 지우기 (CString -> int)
			str_data7.Format(L"%d", int_data7);
			int_data8 = _ttoi(data8);   // 엑셀 숫자 자릿점 지우기 (CString -> int)
			str_data8.Format(L"%d", int_data8);
			int_data9 = _ttoi(data9);   // 엑셀 숫자 자릿점 지우기 (CString -> int)
			str_data9.Format(L"%d", int_data9);
			int_data10 = _ttoi(data10);   // 엑셀 숫자 자릿점 지우기 (CString -> int)
			str_data10.Format(L"%d", int_data10);

			m_ListCtrl1.SetItemText(i-2, 4, data2);  // 1-1
			m_ListCtrl1.SetItemText(i-2, 5, data3);  // 1-2
			m_ListCtrl1.SetItemText(i-2, 6, data4);  // 2-1
			m_ListCtrl1.SetItemText(i-2, 7, data2);  // 2-2
			m_ListCtrl1.SetItemText(i-2, 8, data3);  // 3-1
			m_ListCtrl1.SetItemText(i-2, 9, data4);  // 3-2
		}
	}
	XL.ReleaseExcel();


}


void Cmfc069excelDlg::OnBnClickedExcelWriteBtn()
{
	int m_iMax;
	int nColumncount= m_ListCtrl1.GetItemCount();
	int columnNum = 0;		
	CString m_strFileName;

	m_iMax = nColumncount;
	
	CXLEzAutomation XL(FALSE); // 엑셀 API 함수를 사용하기 위한 클래스 변수 선언
	XL.SetCellValue(++columnNum, 1, _T("")); // SetCellValue:
	XL.SetCellValue(++columnNum, 1, _T("정식과목명"));
	XL.SetCellValue(++columnNum, 1, _T("단축과목명"));
	XL.SetCellValue(++columnNum, 1, _T("교사명"));

	XL.SetCellValue(++columnNum, 1, _T("1-1반"));
	XL.SetCellValue(++columnNum, 1, _T("1-2반"));
	XL.SetCellValue(++columnNum, 1, _T("2-1반")); 
	XL.SetCellValue(++columnNum, 1, _T("2-2반"));
	XL.SetCellValue(++columnNum, 1, _T("3-1반"));
	XL.SetCellValue(++columnNum, 1, _T("3-2반"));
	

	for (int i = 1; i <= m_iMax; i++)
	{
		XL.SetCellValue(1, i + 1, m_ListCtrl1.GetItemText(i-1, 0)); // SetCellValue:
		for (int j = 1; j <= columnNum; j++)
			XL.SetCellValue(j + 1, i + 1, m_ListCtrl1.GetItemText(i-1, j));
	}

	CFileDialog dlg(false, _T("xlsx"), m_strFileName + _T(".xlsx"),
		             OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT | OFN_NOCHANGEDIR,
		             _T("xlsx 파일(*.xls)|*.xls|)"), NULL);
	// OFN_HIDEREADONLY: 읽기 전용 확인란을 숨깁니다.
	// OFN_OVERWRITEPROMPT: 선택한 파일이 이미 있는 경우 다른 이름으로 저장 대화 상자에서 메 사용자는 파일을 덮어쓸 것인지 여부를 확인해야 합니다.
	//
	// OFN_NOCHANGEDIR: 사용자가 파일을 검색하는 동안 디렉터리를 변경한 경우 현재 디렉터리를
	if (dlg.DoModal() == IDOK)
		XL.SaveFileAs(dlg.GetPathName()); // SaveFileAs: 엑셀 파일 저장
	AfxMessageBox(_T("엑셀파일로 저장되었습니다"));
	XL.ReleaseExcel(); // 엑셀 파일 종료
}
