import os
import ctypes
from ctypes import wintypes

def select_items(select_mode: Literal['folder', 'file', 'multi-files'], buffer_size: int = 8192) -> List[str]:
    """
    Opens a Windows dialog to select files or folders.

    Args:
        select_mode (str): The mode of selection. Options: [ 'folder' , 'file' , 'multi-files' ]
        buffer_size (int): The size of the buffer for storing file or folder paths. Defaults to 8192.

    Returns:
        List[str]: A list of selected items (files or folders). Returns an empty list if canceled or an error occurs.

    Raises:
        ValueError: If the select_mode is invalid.
    """
    # 定義 BROWSEINFO 結構體，用於文件夾選擇對話框
    class BROWSEINFO(ctypes.Structure):
        _fields_ = [
            ("hwndOwner", wintypes.HWND),               # 父視窗的句柄，None 表示無父視窗
            ("pidlRoot", wintypes.LPCVOID),             # 根目錄的 PIDL
            ("pszDisplayName", wintypes.LPWSTR),        # 用於接收文件夾名稱的緩衝區
            ("lpszTitle", wintypes.LPCWSTR),            # 對話框標題
            ("ulFlags", wintypes.UINT),                 # 對話框標誌
            ("lpfn", wintypes.LPVOID),                  # 鉤子函數的指標
            ("lParam", wintypes.LPARAM),                # 自定義數據
            ("iImage", wintypes.INT),                   # 文件夾圖標的索引
        ]
    
    # 定義 OPENFILENAME 結構體，用於文件選擇對話框
    class OPENFILENAME(ctypes.Structure):
        _fields_ = [
            ("lStructSize", wintypes.DWORD),            # 結構體的大小（以字節為單位）
            ("hwndOwner", wintypes.HWND),               # 父視窗的句柄，None 表示無父視窗
            ("hInstance", wintypes.HINSTANCE),          # 應用程序實例句柄，通常為 None
            ("lpstrFilter", wintypes.LPCWSTR),          # 文件過濾器字串，用於指定可選的文件類型
            ("lpstrCustomFilter", wintypes.LPWSTR),     # 自定義過濾器
            ("nMaxCustFilter", wintypes.DWORD),         # 自定義過濾器的大小
            ("nFilterIndex", wintypes.DWORD),           # 當前選中的過濾器索引
            ("lpstrFile", wintypes.LPWSTR),             # 文件名緩衝區的指標，用於接收用戶選中的文件路徑
            ("nMaxFile", wintypes.DWORD),               # 文件名緩衝區的大小
            ("lpstrFileTitle", wintypes.LPWSTR),        # 文件名的緩衝區指標（不含路徑）
            ("nMaxFileTitle", wintypes.DWORD),          # 文件名緩衝區的大小
            ("lpstrInitialDir", wintypes.LPCWSTR),      # 初始化對話框的目錄
            ("lpstrTitle", wintypes.LPCWSTR),           # 對話框的標題
            ("Flags", wintypes.DWORD),                  # 對話框的選項標誌
            ("nFileOffset", wintypes.WORD),             # 文件名相對於完整路徑的偏移量
            ("nFileExtension", wintypes.WORD),          # 文件擴展名在路徑中的偏移量
            ("lpstrDefExt", wintypes.LPCWSTR),          # 預設的文件擴展名
            ("lCustData", wintypes.LPARAM),             # 用戶數據，用於自定義回調
            ("lpfnHook", wintypes.LPVOID),              # 鉤子函數的指標
            ("lpTemplateName", wintypes.LPCWSTR),       # 模板名稱，用於自定義對話框界面
            ("pvReserved", wintypes.LPVOID),            # 保留字段，應設為 None
            ("dwReserved", wintypes.DWORD),             # 保留字段，應設為 0
            ("FlagsEx", wintypes.DWORD),                # 擴展標誌
        ]

    # 常量定義
    FOS_PICKFOLDERS = 0x00000020        # 文件夾選擇模式
    OFN_ALLOWMULTISELECT = 0x00000200   # 允許多選
    OFN_EXPLORER = 0x00080000           # 使用新樣式
    OFN_FILEMUSTEXIST = 0x00000008      # 文件必須存在

    # 創建緩衝區，用於接收文件或文件夾路徑
    buffer = ctypes.create_unicode_buffer(buffer_size)

    # 選擇模式: 文件夾
    if select_mode == 'folder':
        shell32 = ctypes.windll.shell32
        ole32 = ctypes.windll.ole32

        # 初始化 COM 庫
        ole32.CoInitialize(None)

        # 定義相關函數
        SHBrowseForFolder = shell32.SHBrowseForFolderW
        SHBrowseForFolder.restype = ctypes.c_void_p  # 設置返回值類型
        SHGetPathFromIDList = shell32.SHGetPathFromIDListW
        CoTaskMemFree = ole32.CoTaskMemFree

        # 初始化 BROWSEINFO 結構體
        bi = BROWSEINFO()
        bi.hwndOwner = None
        bi.pidlRoot = None
        bi.pszDisplayName = ctypes.cast(ctypes.addressof(buffer), wintypes.LPWSTR)
        bi.lpszTitle = "選擇資料夾"
        bi.ulFlags = FOS_PICKFOLDERS
        bi.lpfn = None
        bi.lParam = 0

        # 打開文件夾選擇對話框
        pidl = SHBrowseForFolder(ctypes.byref(bi))
        # 如果選擇成功，解析路徑
        if pidl:
            pidl_ptr = ctypes.cast(pidl, wintypes.LPCVOID)
            SHGetPathFromIDList(pidl_ptr, buffer)
            result = buffer[:].split("\0")
            # 釋放內存
            CoTaskMemFree(pidl_ptr)
            # 釋放 COM 庫
            ole32.CoUninitialize()
            return [result[0]]
        else:
            # 如果取消，釋放 COM 庫
            ole32.CoUninitialize()
            return []
    else:
        # 使用 comdlg32.dll（提供文件選擇對話框的功能）
        comdlg32 = ctypes.windll.comdlg32

        GetOpenFileName = comdlg32.GetOpenFileNameW
        CommDlgExtendedError = comdlg32.CommDlgExtendedError

        # 定義文件過濾器（篩選出可選的文件類型）
        filter_text = "所有檔案\0*.*\0CSV 檔案\0*.csv\0文字檔案\0*.txt\0\0"

        # 初始化 OPENFILENAME 結構體
        ofn = OPENFILENAME()
        # 設置結構體大小
        ofn.lStructSize = ctypes.sizeof(OPENFILENAME)
        # 無父視窗
        ofn.hwndOwner = None
        # 文件過濾器
        ofn.lpstrFilter = filter_text
        # 指定文件緩衝區
        ofn.lpstrFile = ctypes.cast(ctypes.addressof(buffer), wintypes.LPWSTR)
        # 設定緩衝區大小
        ofn.nMaxFile = buffer_size

        # 選擇模式: 文件
        if select_mode == 'file':
            # 對話框標題
            ofn.lpstrTitle = "選擇檔案"
            # 設置標誌
            ofn.Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST
        # 選擇模式: 複數文件
        elif select_mode == 'multi-files':
            # 對話框標題
            ofn.lpstrTitle = "選擇檔案(可複選)"
            # 設置標誌
            ofn.Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_ALLOWMULTISELECT
        else:
            raise ValueError("Invalid select_mode value. Must be 'folder', 'file', or 'multi-files'.")

        # 呼叫文件選擇對話框
        flag = GetOpenFileName(ctypes.byref(ofn))
        # 如果用戶選擇文件
        if flag:
            files = []
            # 獲取緩衝區內容
            result = [item for item in buffer[:].split("\0") if item]
            # 如果選擇多個文件
            if len(result) > 1:
                # 第一個項目是資料夾路徑
                folder = result[0]
                # 拼接路徑
                files = [os.path.join(folder, f) for f in result[1:] if f]
            # 如果僅選擇一個文件
            else:
                files = [result[0]]
            # 返回文件路徑清單
            return files
        # 如果用戶取消或發生錯誤
        else:
            # 如果取消或發生錯誤，處理錯誤碼
            error_code = CommDlgExtendedError()
            if error_code == 0:
                print("Dialog canceled.")
            else:
                print(f"Error occurred. CommDlgExtendedError code: {error_code}")
            return []
