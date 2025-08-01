# coding = utf-8

"""匯入所需到函式庫"""
import datetime
import json
import logging
import os
import shutil
import tkinter as tk
from datetime import timedelta, timezone
from tkinter import messagebox, scrolledtext
from calendar_widget import Calendar
import gspread
import numpy as np
import pandas as pd
import pytz
import requests
from cryptography.fernet import Fernet
from oauth2client.service_account import ServiceAccountCredentials
from pandas import json_normalize
from tkcalendar import Calendar
from datetime import datetime
from tkinter import font as tkfont

"""全域配置指定路徑"""
config_file_path = 'config.json'
temp_directory = 'temp'

if os.path.exists(temp_directory):
    shutil.rmtree(temp_directory)

"""定義日記檔"""
def setup_logging(temp_directory, log_file_path, error_log_path):
    # 每一次都會新增一個temp存放log
    if not os.path.exists(temp_directory):
        os.makedirs(temp_directory)

    # 配置全局日誌
    logging.basicConfig(filename=log_file_path, level=logging.INFO, format='%(asctime)s %(levelname)s:%(message)s', encoding='utf-8')

    # 配置檢查的日誌
    error_logger = logging.getLogger('errorLogger')
    error_handler = logging.FileHandler(error_log_path, encoding='utf-8')
    error_handler.setLevel(logging.INFO)
    error_formatter = logging.Formatter('%(asctime)s %(levelname)s:%(message)s')
    error_handler.setFormatter(error_formatter)
    error_logger.addHandler(error_handler)

    return error_logger

"""配置文件檔設定"""
def load_config(config_file_path):
    try:
        with open(config_file_path, 'r') as file:
            return json.load(file)
    except Exception as e:
        logging.error(f"❌ 無法加載配置文件 {config_file_path}: {e}")
        return None

"""UI介面設定(用來更改config檔日期)"""
def create_date_selector(root, config_file_path):
    date_selector_window = tk.Toplevel(root)
    date_selector_window.title('選擇日期')
    date_selector_window.geometry('300x700')
    date_selector_window.resizable(True, True)
    date_selector_window.configure(bg='lightblue')  # 設定底色

    # 起始日期選擇
    tk.Label(date_selector_window, text='開始日期', font=('Microsoft JhengHei', 12, 'bold'), bg='lightblue').pack(pady=(10, 0))
    start_calendar = Calendar(
        date_selector_window,
        selectmode='day',
        date_pattern='yyyy-mm-dd',
        background='lightblue',
        foreground='black',
        selectbackground='blue',
        selectforeground='red',
        disabledbackground='grey',
        disabledforeground='lightgrey',
        font=('Microsoft JhengHei', 12),
        showweeknumbers=False,
        firstweekday='sunday'
    )
    start_calendar.pack(pady=(0, 10))

    # 結束日期選擇
    tk.Label(date_selector_window, text='結束日期', font=('Microsoft JhengHei', 12, 'bold'), bg='lightblue').pack(pady=(10, 0))
    end_calendar = Calendar(
        date_selector_window,
        selectmode='day',
        date_pattern='yyyy-mm-dd',
        background='lightblue',
        foreground='black',
        selectbackground='blue',
        selectforeground='red',
        disabledbackground='grey',
        disabledforeground='lightgrey',
        font=('Microsoft JhengHei', 12),
        showweeknumbers=False,
        firstweekday='sunday'
    )
    end_calendar.pack(pady=(0, 20))

    # 儲存並產生 CSV
    def save_dates():
        try:
            start_date = start_calendar.get_date()
            end_date = end_calendar.get_date()

            if end_date < start_date:
                messagebox.showwarning("日期錯誤", "結束日期不可早於開始日期")
                return

            # 讀取配置檔案
            with open(config_file_path, 'r', encoding='utf-8') as file:
                config = json.load(file)

            # 更新日期區間
            start_datetime = datetime.strptime(start_date, "%Y-%m-%d")
            end_datetime = datetime.strptime(end_date, "%Y-%m-%d")

            config["time_range"]["start_time"] = start_datetime.strftime("%Y-%m-%d 00:00:00")
            config["time_range"]["end_time"] = end_datetime.strftime("%Y-%m-%d 02:00:00")

            # 寫入回配置檔案
            with open(config_file_path, 'w', encoding='utf-8') as file:
                json.dump(config, file, indent=4, ensure_ascii=False)

            messagebox.showinfo("成功", "日期已成功保存，請不要關閉主視窗!\n點選確定以繼續！")

            date_selector_window.destroy()

            # 呼叫外部函式
            export_data(root)

        except Exception as e:
            messagebox.showerror("錯誤", f"保存日期時發生錯誤：\n{e}")

    save_button_style = {
        "font": ("Microsoft JhengHei", 12, "bold"),
        "bg": "Pink",           # 綠色背景
        "fg": "Black",             # 白色字體
        "activebackground": "Yellow",  # 點擊時的背景色
        "width": 21,
        "height": 2,
        "relief": "raised",        # 浮凸效果
        "bd": 2                    # 邊框寬度
    }

    tk.Button(date_selector_window, text='保存日期，並開始產生 CSV',
            command=save_dates, **save_button_style).pack(pady=10)

# 月報區間選擇
def monthly_report(root, config_file_path):
    entry_window = tk.Toplevel(root)
    entry_window.title("Report Month")
    entry_window.geometry('300x200')
    entry_window.resizable(False, False)
    entry_window.configure(bg='lightblue') 

    def get_date_range():
        try:
            month_input = entry.get()
            year = int(month_input[:4])
            month = int(month_input[4:])
            start_date = datetime(year, month, 1)
            end_date = start_date.replace(day=1, month=start_date.month % 12 + 1) - timedelta(days=1)

            with open(config_file_path, 'r') as file:
                config = json.load(file)

            # 更新配置文件
            config["monthly_report_month"]["start_time"] = start_date.strftime("%Y-%m-%d 00:00:00")
            config["monthly_report_month"]["end_time"] = (end_date + timedelta(days=1)).strftime("%Y-%m-%d 02:00:00")

            with open(config_file_path, 'w') as file:
                json.dump(config, file, indent=4)

            messagebox.showinfo("成功", f"日報區間為 {start_date.strftime('%Y-%m-%d')} 至 {end_date.strftime('%Y-%m-%d')}.\n請稍等")

            entry_window.destroy()

            monthly_report_export(root,month_input)

        except ValueError:
            messagebox.showerror("錯誤", "請輸入 YYYYMM 格式.")

    tk.Label(entry_window, text='輸入月份 (YYYYMM):', font=('Microsoft JhengHei', 12, 'bold'), bg='lightblue').pack()
    entry = tk.Entry(entry_window, width=15, font=('Microsoft JhengHei', 18))
    entry.pack(padx=20, pady=5)

    save_button_style = {
        "font": ("Microsoft JhengHei", 12, "bold"),
        "bg": "Pink",           # 綠色背景
        "fg": "Black",             # 白色字體
        "activebackground": "Yellow",  # 點擊時的背景色
        "width": 20,
        "height": 2,
        "relief": "raised",        # 浮凸效果
        "bd": 2                    # 邊框寬度
    }

    tk.Button(entry_window, text='保存區間，並開始產生 CSV',
            command=get_date_range, **save_button_style).pack(pady=5)

# 月報製作
def monthly_report_export(root,month_input):

    # 日誌路徑
    temp_directory = "temp"
    log_file_path = "temp/monthly_logfile.log"
    error_log_path = "temp/monthly_error_log.log"

    # 設置日誌
    error_logger = setup_logging(temp_directory, log_file_path, error_log_path)

    try:
        # 加載配置
        config = load_config('.\config.json')

        if config:
            # 提取 ClickUp 相關配置
            team_id = config["clickup"]["team_id"]
            user_ids = [user["id"] for user in config["clickup"]["user_ids"]]
            user_name = [user["name"] for user in config["clickup"]["user_ids"]]
            
            start_unix_timestamp = convert_taiwan_time_str_to_unix_timestamp(config["monthly_report_month"]["start_time"])
            end_unix_timestamp = convert_taiwan_time_str_to_unix_timestamp(config["monthly_report_month"]["end_time"])

            # 初始化一个空的 DataFrame，存放clickup api資料
            all_time_entries_df = pd.DataFrame()
            for user_id in user_ids:
                user_time_entries = get_time_entries(team_id, user_id, decrypted_config, start_unix_timestamp, end_unix_timestamp)

                if user_time_entries:
                    user_time_entries_df = pd.DataFrame(user_time_entries)
                    all_time_entries_df = pd.concat([all_time_entries_df, user_time_entries_df], ignore_index=True)

            original_all_time_entries_df = all_time_entries_df.copy()
            for col in original_all_time_entries_df.columns:
                all_time_entries_df = expand_dict_col(all_time_entries_df, col)

            logging.info("✅ 已成功處理 ClickUp 資料")

            # 提取 Google Sheets 相關配置
            credentials_info = config['googlesheet']
            spreadsheet_url = config['spreadsheet_url']['url']
            worksheet_index = config['spreadsheet_url'].get('worksheet_index', 0)
            # 從 Google Sheets 獲取資料併轉換成 DataFrame
            project_info = fetch_data_from_google_sheets(credentials_info, spreadsheet_url, worksheet_index)
            logging.info("✅ 已成功處理 Google Sheets 資料")

            # 對資料欄位和型態進行轉換
            project_info.rename(columns={'List ID (KEY)': 'CU_LIST_ID'}, inplace=True)
            all_time_entries_df.rename(columns={'task_location_list_id': 'List_Id'}, inplace=True)
            project_info['CU_LIST_ID'] = pd.to_numeric(project_info['CU_LIST_ID'], errors='coerce')
            all_time_entries_df['List_Id'] = pd.to_numeric(all_time_entries_df['List_Id'], errors='coerce')
            logging.info("✅ 已成功處理 DataFrame 列名重命名和資料型態轉換")

            # 使用函示對資料進行合併
            js_pjc = process_data(all_time_entries_df, project_info)
            logging.info("✅ 已成功處理 DataFrame 合併")

            # 使用函示新增欄位
            js_pjc['flattened_tags'] = js_pjc.apply(flatten_tags, axis=1)
            js_pjc['duration_overtime'] = js_pjc.apply(lambda x: x['duration'] if contains_overtime(x['tags']) else 0, axis=1)
            js_pjc.loc[js_pjc['tags'].apply(contains_overtime), 'duration'] = 0
            js_pjc['extracted_numbers'] = js_pjc['tags'].apply(extract_numbers)
            js_pjc['extracted_numbers'] = js_pjc['extracted_numbers'].str[0]
            logging.info("✅ 已成功處理 DataFrame 新增欄位")

            # 選擇需要的列
            columns_to_select = ['user_id', 'user_username', 'start', '系統代碼', '*專案代碼','工作代碼', 'extracted_numbers', 'duration', 'duration_overtime','task_name', 'CU Folder']
            js_pjc_select = js_pjc[columns_to_select].copy()
            logging.info("✅ 已成功選擇特定列並創建新的 DataFrame")

            # 重新命名欄位
            js_pjc_select.rename(columns={'extracted_numbers': '工作別'}, inplace=True)
            js_pjc_select.rename(columns={'duration': '正常工時'}, inplace=True)
            js_pjc_select.rename(columns={'duration_overtime': '加班工時'}, inplace=True)
            js_pjc_select.rename(columns={'task_name': '任務'}, inplace=True)
            js_pjc_select.rename(columns={'CU Folder': 'CU專案'}, inplace=True)
            logging.info("✅ 已成功重新命名欄位")

            # 使用函式轉換欄位成年月日格式
            timestamp_columns = ['start']
            js_pjc_select = convert_multiple_timestamps(js_pjc_select, timestamp_columns)
            logging.info("✅ 已成功轉換欄位成年月日格式")

            # 使用函式轉換欄位成小時單位
            columns_to_convert = ['加班工時', '正常工時']
            js_pjc_select = convert_ms_to_hours(js_pjc_select, columns_to_convert)
            js_pjc_select['總工時'] = js_pjc_select['加班工時'] + js_pjc_select['正常工時']
            logging.info("✅ 已成功轉換欄位成小時單位")

    except Exception as e:
        logging.error(f"❌ 資料處理過程中發生錯誤：{e}")

    try:
        output_path = f"./OUTPUT/合規組_{month_input}_月報表_FCS.xlsx"
        os.makedirs("./OUTPUT", exist_ok=True)

        # 開啟統一 ExcelWriter
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            workbook = writer.book

            grouped_by_user = js_pjc_select.groupby('user_username')

            for username, user_group in grouped_by_user:
                try:
                    # 擷取檔名
                    first_name = username.split('.')[0] if '.' in username else username.split(' ')[0]

                    # 建 summary_rows（如前面邏輯）
                    summary_rows = []
                    grouped_by_project = user_group.groupby('CU專案')
                    for project, proj_df in grouped_by_project:
                        project_total = proj_df['總工時'].sum()
                        summary_rows.append([project, round(project_total, 2)])

                        grouped_by_task = proj_df.groupby('任務')
                        for task, task_df in grouped_by_task:
                            task_total = task_df['總工時'].sum()
                            summary_rows.append(['　' + str(task), round(task_total, 2)])

                    # 插入標題、欄名
                    header_rows = [[f"{month_input} 工作月報表", ""], ["專案名稱/工作內容", "時數"]]
                    summary_df = pd.DataFrame(header_rows + summary_rows)

                    # ✅ 加入總計列
                    project_rows = [row for row in summary_rows if isinstance(row[0], str) and not row[0].startswith('　')]
                    total_hours = round(sum(row[1] for row in project_rows), 2)
                    summary_df.loc[len(summary_df.index)] = ['總計', total_hours]

                    # 建立工作表
                    summary_df.to_excel(writer, index=False, header=False, sheet_name=first_name)
                    worksheet = writer.sheets[first_name]

                    # 設定欄寬
                    worksheet.set_column('A:A', 70)
                    worksheet.set_column('B:B', 7)

                    # 寫入每格與格式（與前面提供一致）
                    for row_idx in range(len(summary_df)):
                        val_col0 = summary_df.iloc[row_idx, 0]
                        val_col1 = summary_df.iloc[row_idx, 1]

                        is_header = row_idx < 2
                        is_project_row = isinstance(val_col0, str) and not val_col0.startswith('　')
                        is_total_row = row_idx == len(summary_df) - 1

                        if is_header:
                            fmt_col0 = workbook.add_format({
                                'font_name': 'DFKai-SB', 'bold': True, 'align': 'center',
                                'valign': 'vcenter', 'border': 1
                            })
                            fmt_col1 = fmt_col0
                        elif is_project_row or is_total_row:
                            fmt_col0 = workbook.add_format({
                                'font_name': 'DFKai-SB', 'bold': True,
                                'border': 1, 'bg_color': '#D9E1F2'
                            })
                            fmt_col1 = workbook.add_format({
                                'font_name': 'DFKai-SB', 'bold': True,
                                'border': 1, 'bg_color': '#D9E1F2' #, 'num_format': '0.0'
                            })
                        else:
                            fmt_col0 = workbook.add_format({
                                'font_name': 'DFKai-SB', 'border': 1
                            })
                            fmt_col1 = workbook.add_format({
                                'font_name': 'DFKai-SB', 'border': 1 #, 'num_format': '0.0'
                            })

                        worksheet.write(row_idx, 0, val_col0, fmt_col0)
                        worksheet.write(row_idx, 1, val_col1, fmt_col1)

                    logging.info(f"✅ 加入 {username} 工作表")

                except Exception as e:
                    logging.error(f"❌ 處理使用者 {username} 時錯誤：{e}")

        logging.info(f"📁 所有使用者工作月報表已成功匯出。")

    except Exception as e:
        logging.error(f"❌ 檔案匯出時發生錯誤：{e}")
        logging.info("✅ 所有動作已完成，請至OUTPUT查看檔案")
    finally:
        try:
            root.destroy()  # 關閉主視窗
            messagebox.showinfo("完成", "已完成匯出")  # 彈出提示訊息
        except Exception as e:
            logging.warning(f"❌ 無法關閉視窗或顯示完成訊息：{e}")

    root.destroy()  # 關閉主視窗
    messagebox.showinfo("完成", "已完成匯出") 

# 開啟Log
def open_log_window(root, log_file_path):
    log_window = tk.Toplevel(root)
    log_window.title("Log Window")

    log_text = tk.Text(log_window, width=300, height=160)
    log_text.pack()

    try:
        # Specify the encoding here
        with open(log_file_path, 'r', encoding='utf-8') as file:
            log_text.insert(tk.END, file.read())
    except FileNotFoundError:
        log_text.insert(tk.END, "Log file not found.")
    except UnicodeDecodeError:
        log_text.insert(tk.END, "Error reading the log file: encoding issue.")

"""金鑰加密處理"""
# 生成的密鑰
key = b'h8ismKmlffQ56ReTOBLFTPkU8PlVfjyff2EwgQKrHTE='
# Fernet 密鑰算法
cipher_suite = Fernet(key)
# 已加密數據
encrypted_data = b'gAAAAABlpfX8WV1yG2gyTkW1AGNeNCpBd4iCeG6hjlMPVmBnJbnmVc5cZ0OUYcAEmgFii7AX_pyH8zvQ4UsZXz8bzzEJsQ421LDkiAcmmZ9Mg5FauwUvGWYMZM_DkOC1hJNMuOtYjru0'
# 解密數據
decrypted_config = cipher_suite.decrypt(encrypted_data)

"""配置文件時間年月日轉Unix"""
def convert_taiwan_time_str_to_unix_timestamp(time_str, time_format="%Y-%m-%d %H:%M:%S"):
    """將台灣時間字串轉換 Unix 時間戳（毫秒）"""
    #解析時間
    taiwan_timezone = pytz.timezone("Asia/Taipei")
    local_time = datetime.strptime(time_str, time_format)
    local_time = taiwan_timezone.localize(local_time)
    utc_time = local_time.astimezone(pytz.utc)
    # 轉換為 Unix
    unix_timestamp = int(utc_time.timestamp() * 1000)
    return unix_timestamp

"""clickup api 資料爬取"""
def get_time_entries(team_id, user_id, decrypted_config, start_unix_timestamp, end_unix_timestamp):
    """發起請求獲取指定用戶資料"""
    try:
        url = f"https://api.clickup.com/api/v2/team/{team_id}/time_entries"
        headers = {
            "Content-Type": "application/json",
            "Authorization": decrypted_config
        }
        query_params = {
            "assignee": str(user_id),
            "start_date": start_unix_timestamp,
            "end_date": end_unix_timestamp
        }
        response = requests.get(url, headers=headers, params=query_params)

        if response.status_code == 200:
            logging.info(f"✅ 成功取得用戶 {user_id} ")
            return response.json().get('data', [])
        else:
            logging.warning(f"❌ 用戶 {user_id} 的資料獲取失敗: {response.text}")
            return []

    except Exception as e:
        logging.error(f"❌ 在取得用戶 {user_id} 的資料時發生錯誤: {e}")
        return []

"""google sheet api 資料爬取"""
def fetch_data_from_google_sheets(credentials_info, spreadsheet_url, worksheet_index=0):
    # 設置 Google Sheets API 認證
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_info, scope)
    gc = gspread.authorize(credentials)

    # 打開指定的 Google Sheets 工作表
    workbook = gc.open_by_url(spreadsheet_url)
    worksheet = workbook.get_worksheet(worksheet_index)

    # 將資料轉成DataFrame
    data = worksheet.get_all_values()
    df = pd.DataFrame(data)
    df.columns = df.iloc[0]
    df = df[1:]
    return df

"""自定義函數"""

"""定義合併googlesheet和clickup函式"""
def process_data(all_entries, df):
    if 'List_Id' in all_entries.columns and 'CU_LIST_ID' in df.columns:
        processed_data = pd.merge(all_entries, df, left_on='List_Id', right_on='CU_LIST_ID', how='left')
        logging.info("✅ clickup_timesheet 和 project_info 資料合併成功")
    else:
        logging.error("❌ 合併失敗：缺少必要的列")
        processed_data = pd.DataFrame()
    return processed_data

"""定義展開字典類型的列函式"""
def expand_dict_col(df, column):
    if isinstance(df[column].iloc[0], dict):
        expanded_cols = df[column].apply(pd.Series)
        expanded_cols = expanded_cols.rename(lambda x: f"{column}_{x}", axis='columns')
        df = pd.concat([df.drop(column, axis=1), expanded_cols], axis=1)
    return df

"""定義拆解tag欄位函式"""
def flatten_tags(row):
    tags = row['tags']
    if isinstance(tags, str):
        tags = json.loads(tags.replace("'", '"'))
    return [tag['name'] for tag in tags if 'name' in tag]

"""定義檢查'overtime'函式"""
def contains_overtime(tag_list):
    return any(tag.get('name') == 'overtime' for tag in tag_list)

"""定義對tags字典欄位條件取值函式"""
def extract_numbers(tags):
    numbers = []
    try:
        if tags and isinstance(tags, list):
            for item in tags:
                if isinstance(item, dict) and 'name' in item:
                    name = item['name']
                    number = name.split('_')[0] if '_' in name else None
                    if number:
                        numbers.append(number)
        logging.debug(f"從name提取數字: {numbers}")
        return numbers
    except Exception as e:
        logging.error(f"❌ 在提取數字過程中發生錯誤: {e}")
        return []

"""定義日期轉換成年月日呈現函式"""
def convert_multiple_timestamps(df, timestamp_columns, timezone='Asia/Taipei'):
    try:
        tz = pytz.timezone(timezone)
        for column in timestamp_columns:
            # 转换每个指定的时间戳列
            df[column] = pd.to_datetime(df[column], unit='ms', utc=True)
            df[column] = df[column].dt.tz_convert(tz).dt.date
        return df
    except Exception as e:
        print(f"❌ 轉換時間戳記發生錯誤: {e}")
        return None

"""定義毫秒轉換成小時函式"""
def convert_ms_to_hours(df, columns):
    ms_to_hours = 3600000
    for column in columns:
        try:
            df[column] = pd.to_numeric(df[column], errors='coerce')
            df[column] = (df[column] / ms_to_hours).apply(lambda x: np.floor(x * 10) / 10)
        except KeyError:
            print(f"'{column}'不存在在資料集")
        except Exception as e:
            print(f"轉換小時單位時發生錯誤 '{column}': {e}")
    return df

"""定義請假登打檢查函式"""
def check_leave_work_type(df, work_code, work_type, error_logger):
    try:
        for start, group_df in df.groupby('start'):
            filtered_df = group_df[(group_df['工作代碼'] == work_code) & (group_df['工作別'] != work_type)]
            is_error_found = False
            error_message = ""

            for index, row in filtered_df.iterrows():
                is_error_found = True
                error_logger.error(f"🔴 {row['start']}, {row['user_username']} 的請假工作別有問題\n")

            if not is_error_found:
                error_logger.info(f"🟢 {start} 沒有發現任何工作代碼為 '{work_code}' 且工作別不等於 {work_type} 的情況")

    except Exception as e:
        error_logger.error(f"❌ 分組和篩選過程中發生錯誤: {e}")

"""定義請假工作別檢查函式"""
def check_leave_work_type(df, work_code, work_type, error_logger):
    try:
        # 確保 DataFrame 中的 '工作代碼' 和 '工作別' 欄位是字符串格式
        # 如果原本不是字符串，則進行轉換
        df['工作代碼'] = df['工作代碼'].astype(str)
        df['工作別'] = df['工作別'].astype(str)

        # 同樣確保傳入函式的 work_code 和 work_type 參數是字符串
        work_code = str(work_code)
        work_type = str(work_type)

        # 遍歷按 'start' 分組的 DataFrame
        for start, group_df in df.groupby('start'):
            # 篩選出符合特定工作代碼且工作別不符合條件的數據
            filtered_df = group_df[(group_df['工作代碼'] == work_code) & (group_df['工作別'] != work_type)]
            is_error_found = False
            error_message = ""

            # 檢查篩選出的數據，並構建錯誤信息
            for index, row in filtered_df.iterrows():
                is_error_found = True
                error_logger.error(f"{'🔴'} {row['start']}, {row['user_username']} 的請假工作別有問題\n")

            # 根據是否發現錯誤記錄不同的日誌信息
            if not is_error_found:
                error_logger.info(f"🟢 {start} 沒有發現任何工作代碼為 '{work_code}' 且工作別不等於 {work_type} 的情況")

    except Exception as e:
        # 處理並記錄任何在過程中出現的異常
        error_logger.error(f"分組和篩選過程中發生錯誤: {e}")

"""定義工作別檢查函式"""
def check_work_category(df, valid_categories, error_logger):
    try:
        for start, group_df in df.groupby('start'):
            filtered_category = group_df[~group_df['工作別'].isin(valid_categories)]

            is_error_found = False
            error_message = ""

            for index, row in filtered_category.iterrows():
                is_error_found = True
                username = row['user_username']
                error_logger.error(f"{'🔴'} {start}, {username} 存在工作別不正確情況\n")

            if not is_error_found:
                error_logger.info(f"🟢 {start} 沒有工作別不正確情況")

    except Exception as e:
        error_logger.error(f"{'❌ '} 分組和篩選 '工作別' 過程中發生錯誤: {e}")

"""定義正常工時超時檢查函式"""
def check_normal_working_hours(df, max_hours, error_logger):

    config = load_config('.\config.json')
    try:
        grouped_by_date = df.groupby('start')
        for start, group in grouped_by_date:
            grouped = group.groupby(['user_username', 'user_id'])['正常工時'].sum()

            is_error_found = False
            error_message = ""

            user_list = config["clickup"]["user_ids"]
            grouped_user_ids = {uid for (_, uid) in grouped.keys()}

            for user in user_list:
                user_id = user["id"]
                user_name = user["name"]
                
                if user_id not in grouped_user_ids:
                    is_error_found = True
                    error_logger.error(f"{'🔴'} {start}, {user_name} 未填工時。")

            for (user, user_id), total_hours in grouped.items():
                if total_hours > max_hours:
                    is_error_found = True
                    error_logger.error(f"{'🔴'} {start}, {user} 的正常工時超過 {total_hours} 小時。")
                elif total_hours < max_hours:
                    is_error_found = True
                    error_logger.error(f"{'🔴'} {start}, {user} 的正常工時未滿 {max_hours} 小時。")

            if not is_error_found:
                error_logger.info(f"🟢 {start} 没有正常工時打錯情况")

    except Exception as e:
        error_logger.error(f"分組和篩選 '正常工時' 過程中發生錯誤: {e}")

"""定義工時單位登打檢查函式"""
def check_time_unit_compliance(df, error_logger):
    try:
        from decimal import ROUND_HALF_UP, Decimal
        df['正常工時'] = df['正常工時'].apply(lambda x: Decimal(x).quantize(Decimal('1'), rounding=ROUND_HALF_UP))
        df['加班工時'] = df['加班工時'].apply(lambda x: Decimal(x).quantize(Decimal('1'), rounding=ROUND_HALF_UP))
        condition_normal = (df['正常工時'] * 60) - 6 * ((df['正常工時'] * 60) // 6) != 0
        condition_overtime = (df['加班工時'] * 60) - 6 * ((df['加班工時'] * 60) // 6) != 0
        condition = condition_normal | condition_overtime

        for date, group_divide in df.groupby('start'):
            found_non_divisible = False
            error_message = ""

            for index, row in group_divide.iterrows():
                if condition_normal[index] or condition_overtime[index]:
                    found_non_divisible = True
                    error_logger.debug(f"{'🔴'} {date}, {row['user_username']} 的工時單位錯誤\n")

            if found_non_divisible:
                logging.warning(f"{'🔴'} {date} 的以下紀錄工時單位錯誤:\n{error_message}")
                
            else:
                error_logger.info(f"🟢 {date} 没有時間單位不符合情况")

    except Exception as e:
        error_logger.error(f"🔴 檢查工時被6整除發生錯誤: {e}")

def export_data(root):

    # 日誌路徑
    temp_directory = "temp"
    log_file_path = "temp/logfile.log"
    error_log_path = "temp/error_log.log"

    # 設置日誌
    error_logger = setup_logging(temp_directory, log_file_path, error_log_path)

    try:
        # 加載配置
        config = load_config('.\config.json')

        if config:
            # 提取 ClickUp 相關配置
            team_id = config["clickup"]["team_id"]
            user_ids = [user["id"] for user in config["clickup"]["user_ids"]]
            user_name = [user["name"] for user in config["clickup"]["user_ids"]]
            
            start_unix_timestamp = convert_taiwan_time_str_to_unix_timestamp(config["time_range"]["start_time"])
            end_unix_timestamp = convert_taiwan_time_str_to_unix_timestamp(config["time_range"]["end_time"])

            # 初始化一个空的 DataFrame，存放clickup api資料
            all_time_entries_df = pd.DataFrame()
            for user_id in user_ids:
                user_time_entries = get_time_entries(team_id, user_id, decrypted_config, start_unix_timestamp, end_unix_timestamp)

                if user_time_entries:
                    user_time_entries_df = pd.DataFrame(user_time_entries)
                    all_time_entries_df = pd.concat([all_time_entries_df, user_time_entries_df], ignore_index=True)

            original_all_time_entries_df = all_time_entries_df.copy()
            for col in original_all_time_entries_df.columns:
                all_time_entries_df = expand_dict_col(all_time_entries_df, col)

            logging.info("✅ 已成功處理 ClickUp 資料")

            # 提取 Google Sheets 相關配置
            credentials_info = config['googlesheet']
            spreadsheet_url = config['spreadsheet_url']['url']
            worksheet_index = config['spreadsheet_url'].get('worksheet_index', 0)
            # 從 Google Sheets 獲取資料併轉換成 DataFrame
            project_info = fetch_data_from_google_sheets(credentials_info, spreadsheet_url, worksheet_index)
            logging.info("✅ 已成功處理 Google Sheets 資料")

            # 對資料欄位和型態進行轉換
            project_info.rename(columns={'List ID (KEY)': 'CU_LIST_ID'}, inplace=True)
            all_time_entries_df.rename(columns={'task_location_list_id': 'List_Id'}, inplace=True)
            project_info['CU_LIST_ID'] = pd.to_numeric(project_info['CU_LIST_ID'], errors='coerce')
            all_time_entries_df['List_Id'] = pd.to_numeric(all_time_entries_df['List_Id'], errors='coerce')
            logging.info("✅ 已成功處理 DataFrame 列名重命名和資料型態轉換")

            # 使用函示對資料進行合併
            js_pjc = process_data(all_time_entries_df, project_info)
            logging.info("✅ 已成功處理 DataFrame 合併")

            # 使用函示新增欄位
            js_pjc['flattened_tags'] = js_pjc.apply(flatten_tags, axis=1)
            js_pjc['duration_overtime'] = js_pjc.apply(lambda x: x['duration'] if contains_overtime(x['tags']) else 0, axis=1)
            js_pjc.loc[js_pjc['tags'].apply(contains_overtime), 'duration'] = 0
            js_pjc['extracted_numbers'] = js_pjc['tags'].apply(extract_numbers)
            js_pjc['extracted_numbers'] = js_pjc['extracted_numbers'].str[0]
            logging.info("✅ 已成功處理 DataFrame 新增欄位")

            # 選擇需要的列
            # columns_to_select = ['user_id', 'user_username', 'start', '系統代碼', '*專案代碼','專案名稱(PJC)/系統名稱','工作代碼', 'extracted_numbers', 'duration', 'duration_overtime','task_name']
            columns_to_select = ['user_id', 'user_username', 'start', '系統代碼', '*專案代碼','工作代碼', 'extracted_numbers', 'duration', 'duration_overtime','task_name']
            js_pjc_select = js_pjc[columns_to_select].copy()
            logging.info("✅ 已成功選擇特定列並創建新的 DataFrame")

            # 使用 .loc 來更新 'task.name' 列
            js_pjc_select.loc[:, 'task_name'] = '(' + js_pjc['task_custom_id'] + ') ' + js_pjc_select['task_name'] + ' : ' + js_pjc['description']
            logging.info("✅ 已更新 'task_name' 列")

            # 重新命名欄位
            js_pjc_select.rename(columns={'extracted_numbers': '工作別'}, inplace=True)
            js_pjc_select.rename(columns={'duration': '正常工時'}, inplace=True)
            js_pjc_select.rename(columns={'duration_overtime': '加班工時'}, inplace=True)
            js_pjc_select.rename(columns={'task_name': '工作說明'}, inplace=True)
            logging.info("✅ 已成功重新命名欄位")

            # 使用函式轉換欄位成年月日格式
            timestamp_columns = ['start']
            js_pjc_select = convert_multiple_timestamps(js_pjc_select, timestamp_columns)
            logging.info("✅ 已成功轉換欄位成年月日格式")

            # 使用函式轉換欄位成小時單位
            columns_to_convert = ['加班工時', '正常工時']
            js_pjc_select = convert_ms_to_hours(js_pjc_select, columns_to_convert)
            logging.info("✅ 已成功轉換欄位成小時單位")

            # 使用函式檢查規則併寫入error_log
            pjc_check = js_pjc_select.copy()
            check_leave_work_type(pjc_check, '00000000', 16, error_logger)
            valid_categories = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19']
            check_work_category(pjc_check, valid_categories, error_logger)
            check_normal_working_hours(pjc_check, 8, error_logger)
            check_time_unit_compliance(pjc_check, error_logger)
            logging.info("✅ 已成功匯出error_log")

    except Exception as e:
        logging.error(f"❌ 資料處理過程中發生錯誤：{e}")

    try:
    #     # 首先按日期分组
        js_pjc_select['start'] = pd.to_datetime(js_pjc_select['start'])
        grouped_by_date = js_pjc_select.groupby(js_pjc_select['start'].dt.date)

        # 為每一個日期建立一個文件夾並指定匯出資料夾
        for date, date_group in grouped_by_date:
            date_format = date.strftime("%Y%m%d")
            date_folder_path = f".\OUTPUT\{date_format}"
            if not os.path.exists(date_folder_path):
                os.makedirs(date_folder_path)
                print(f"日期文件夾已建立：{date_folder_path}")
            else:
                print(f"日期文件夾已建立：{date_folder_path}")

            # 按用戶分组成每一個檔案
            grouped_by_user = date_group.groupby('user_username')
            for username, user_group in grouped_by_user:
                try:
                    # 文件名設定
                    if '.' in username:
                        first_name = username.split('.')[0]
                    else:
                        first_name = username.split(' ')[0]

                    # 'start'列是datetime格式，並格式化為'YYYY-MM-DD'
                    user_group['start'] = user_group['start'].dt.strftime('%Y/%m/%d')

                    # 刪除不必要的列
                    columns_to_drop = ['user_id', 'user_username']
                    user_group.drop(columns=columns_to_drop, inplace=True, errors='ignore')

                    # 建立文件名並保存DataFrame為CSV文件
                    filename = os.path.join(date_folder_path, f"PJC_{first_name}.csv")
                    user_group.to_csv(filename, encoding='big5', index=False, header=False)
                    logging.info(f"📁 儲存dataframe {username} 到 CSV 檔案 {filename}")

                except Exception as e:
                    logging.error(f"❌ 處理用戶名 {username} 時發生錯誤：{e}")

                logging.info(f"✅ 所有文件已成功匯出至 '{date_folder_path}'。")

    except Exception as e:
        logging.error(f"❌ 檔案匯出時發生錯誤：{e}")
        logging.info("✅ 所有動作已完成，請至OUTPUT查看檔案")
    finally:
        try:
            root.destroy()  # 關閉主視窗
            messagebox.showinfo("完成", "已完成匯出")  # 彈出提示訊息
        except Exception as e:
            logging.warning(f"❌ 無法關閉視窗或顯示完成訊息：{e}")

"""主程式"""
def main():

    root = tk.Tk()
    root.title("PJC匯出工具")
    root.geometry("300x200")
    root.configure(bg='lightblue')  # 設定底色

    # 日誌路徑
    log_file_path = "temp/logfile.log"

    # 建立日期選擇器和日誌查看按鈕
    def create_hover_button(parent, text, command, normal_bg="Pink", hover_bg="HotPink"):
        # 建立兩種字型：正常 / 粗體
        normal_font = tkfont.Font(family="Microsoft JhengHei", size=12, weight="normal")
        bold_font = tkfont.Font(family="Microsoft JhengHei", size=12, weight="bold")

        # 建立按鈕
        button = tk.Button(
            parent,
            text=text,
            command=command,
            width=25,
            height=2,
            bg=normal_bg,
            fg="Black",
            font=normal_font,
            relief="raised",
            activebackground="#45a049",
        )
        button.pack(pady=5)

        # 滑入滑出事件：改字體與背景色
        button.bind("<Enter>", lambda e: [button.config(bg=hover_bg, font=bold_font)])
        button.bind("<Leave>", lambda e: [button.config(bg=normal_bg, font=normal_font)])

        return button

    # 建立按鈕
    create_hover_button(root, "匯出PJC", lambda: create_date_selector(root, config_file_path))
    create_hover_button(root, "月報下載", lambda: monthly_report(root, config_file_path))
    create_hover_button(root, "查看日誌", lambda: open_log_window(root, log_file_path))

    root.mainloop()
    
if __name__ == "__main__":
    main()
