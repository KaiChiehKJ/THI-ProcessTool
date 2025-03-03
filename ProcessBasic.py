import pandas as pd
import os 
import shutil

def create_folder(folder_name):
    """建立資料夾"""
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
    return os.path.abspath(folder_name)

def delete_folders(deletelist):
    """
    刪除資料夾
    deletelist(list):需要為皆為路徑的list
    """
    for folder_name in deletelist: 
        if os.path.exists(folder_name): # 檢查資料夾是否存在
            shutil.rmtree(folder_name) # 刪除資料夾及其內容
        else:
            print(f"資料夾 '{folder_name}' 不存在。")

def getdatelist(time1, time2):
    '''
    建立日期清單
    time1、time2(str):為%Y-%M-%D格式的日期字串
    '''
    if time1 > time2:
        starttime = time2
        endtime = time1
    else:
        starttime = time1
        endtime = time2

    date_range = pd.date_range(start=starttime, end=endtime)
    datelist = [d.strftime("%Y%m%d") for d in date_range]
    return datelist

def findfiles(filefolderpath, filetype='.csv', recursive=True):
    """
    尋找指定路徑下指定類型的檔案，並返回檔案路徑列表。

    Args:
        filefolderpath (str): 指定的檔案路徑。
        filetype (str, optional): 要尋找的檔案類型，預設為 '.csv'。
        recursive (bool, optional): 是否檢索所有子資料夾，預設為 True；反之為False，僅查找當前資料夾的所有file。

    Returns:
        list: 包含所有符合條件的檔案路徑的列表。
    """
    filelist = []

    if recursive:
        # 遍歷資料夾及其子資料夾
        for root, _, files in os.walk(filefolderpath):
            for file in files:
                if file.endswith(filetype):
                    file_path = os.path.join(root, file)
                    filelist.append(file_path)
    else:
        # 僅檢索當前資料夾
        for file in os.listdir(filefolderpath):
            file_path = os.path.join(filefolderpath, file)
            if os.path.isfile(file_path) and file.endswith(filetype):
                filelist.append(file_path)

    return filelist

def move_column(df, column_name, insert_index):
    """
    移動DataFrame中的既存欄位到指定位置。

    Args:
        df (pd.DataFrame): 要操作的DataFrame。
        column_name (str): 要移動的欄位名稱。
        insert_index (int): 欲插入的新位置索引（從0開始）。

    Returns:
        pd.DataFrame: 調整後的DataFrame。
    """
    if column_name not in df.columns:
        raise ValueError(f"Column '{column_name}' does not exist in DataFrame.")
    
    columns = df.columns.tolist() # 取得目前欄位順序

    columns.remove(column_name) # 移除該欄位

    columns.insert(insert_index, column_name) # 在指定位置插入該欄位
    
    return df[columns] # 重新排列DataFrame

def get_filename(path, extension=False):
    """
    從檔案路徑中提取檔名，可選擇是否包含副檔名。

    Args:
        path (str): 檔案的完整路徑。
        extension (bool, optional): 是否包含副檔名，預設為 False。

    Returns:
        str: 檔名（根據 extension 參數決定是否包含副檔名）。
    """
    filename = os.path.basename(path)
    if not extension:
        filename = os.path.splitext(filename)[0]
    return filename

def get_excel_sheet_names(path):
    """
    取得 Excel 檔案中的所有工作表名稱。

    Args:
        path (str): Excel 檔案的路徑。

    Returns:
        list: 工作表名稱列表。
    """
    try:
        sheet_names = pd.ExcelFile(path).sheet_names
        return sheet_names
    except FileNotFoundError:
        print(f"檔案不存在：{path}")
        return []
    except Exception as e:
        print(f"發生錯誤：{e}")
        return []

def get_percent_columns(df, columns='Trips'):
    """
    計算百分比欄位，並插入到指定的 columns 欄位後面。

    Args:
        df (DataFrame): 輸入的資料框。
        columns (str): 用來計算百分比的欄位名稱。

    Returns:
        DataFrame: 包含新插入的 Percent 欄位的資料框。
    """
    total_value = df[columns].sum()
    df['Percent'] = (df[columns] / total_value) * 100
    df['Percent'] = df['Percent'].round(2).astype(str) + "%"

    # 找到 columns 欄位的位置，將 Percent 插入在其後面
    col_index = df.columns.get_loc(columns) + 1
    cols = list(df.columns)
    cols.insert(col_index, cols.pop(cols.index('Percent')))
    df = df[cols]

    return df
