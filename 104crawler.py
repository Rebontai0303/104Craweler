'''
author rebontai 20250818
'''

import time
import sv_ttk
import darkdetect
import pandas as pd
import tkinter as tk
import sys
from tkinter import ttk
from tkinter import messagebox
from os import path
from datetime import datetime
from threading import Thread
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from pywinstyles import change_header_color, apply_style
from warnings import filterwarnings
from openpyxl import load_workbook

filterwarnings("ignore")

retry = 0
max_retry = 3

# 建立item list
job_titles = list()          
job_companyname = list()     
job_url = list()             
job_companyindustry = list()                
job_year = list()
job_announcedate = list() 
job_location = list()
job_worktime = list() 
job_edu = list()
job_salary = list()  
job_content = list()         
job_category = list()      
job_dept = list()    
job_specialty = list()      
job_others = list()
url_cache = list()    

data = pd.DataFrame()

def print_message(message, function_name, type):
    '''
    print message.\n
    顯示通知訊息之tk視窗\n
    種類涵蓋\n
    err: 錯誤訊息\n
    info: 一般訊息\n
    war: 警告訊息

    Args:
        message (str): 訊息內容
        function_name (str): 函式名稱
        type (str): 訊息類型
    Return:
        NA.    
    '''
    match type:
        case "err":
            print (f"❗{function_name}發生錯誤: {message}❗")
            log_text.insert("end", f"❗{function_name}發生錯誤: {message}❗\n")
            log_text.see("end")
        case "info":
            print (f"ℹ️{function_name}訊息: {message}ℹ️")
            log_text.insert("end", f"ℹ️{function_name}訊息: {message}ℹ️\n")
            log_text.see("end")
        case "war":
            print (f"⚠️{function_name}警告: {message}⚠️")
            log_text.insert("end", f"⚠️{function_name}警告: {message}⚠️\n")
            log_text.see("end")        
    

def open_main_window():
    '''
    setting tkinter main window.\n
    主要視窗設定\n
    使用主題 - darkdetect\n
    使用字體 - 標楷體(DFKai-SB)

    Args:
        NA.
    Return:
        NA.
    '''
    global root, theme_switch, keyword_pack
    # tkinter視窗設定
    root = tk.Tk()
    # 主題參數
    sv_ttk.set_theme(darkdetect.theme())
    # 視窗標題
    root.title("104人力銀行爬蟲程式")
    # 視窗外觀
    window_width = root.winfo_screenwidth()    # 取得螢幕寬度
    window_height = root.winfo_screenheight()  # 取得螢幕高度
    width = 600
    height = 320
    left = int((window_width - width)/2)       # 計算左上 x 座標
    top = int((window_height - height)/2)      # 計算左上 y 座標
    root.geometry(f"{width}x{height}+{left}+{top}")
    root.resizable(False, False)               # 設定視窗不可調整大小
    
    # tk視窗設定
    keyword_pack = tk.StringVar() 
    ttk.Label(
        root, 
        text="請輸入104人力銀行關鍵字:", 
        font=("DFKai-SB", 20)
    ).pack(pady=25)
    ttk.Entry(
        root, 
        textvariable = keyword_pack, 
        font = ("DFKai-SB", 20), 
        width = 50
    ).pack(pady=10, padx=50)
    ttk.Label(
        root, 
        text = "❗檔案輸出完成將放置於您電腦中的「下載」資料夾中❗", 
        font = ("DFKai-SB", 11)
    ).pack(pady=10)
    ttk.Button(
        root, 
        text="開始", 
        command=on_start
    ).pack()

    # 淺深滑桿設定
    frame = ttk.Frame(root, padding="10").pack(expand=True, fill="both")
    bottom_frame = ttk.Frame(frame)
    bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)

    #copyright
    ttk.Label(bottom_frame, text="Copyright © 2025 Rebontai", font=("Arial", 10)).pack(side=tk.LEFT, padx=10, pady=10)  

    # 元件style設定
    style = ttk.Style()
    style.configure(
        "TButton", 
        font=("DFKai-SB", 20)
    )
    style.configure(
        "Switch.TCheckbutton", 
        font=("DFKai-SB", 10)
    )  
    theme_switch = ttk.Checkbutton(
        bottom_frame, 
        style="Switch.TCheckbutton"
    )
    
    
    # 滑桿綁定當前主題
    theme_switch.pack(side=tk.RIGHT, padx=10, pady=10)
    # 根據主題設定font和label
    if sv_ttk.get_theme() == "dark":
        theme_switch.configure(text="深色模式")
        theme_switch.state(["selected"])
    else:
        theme_switch.configure(text="淺色模式")
        theme_switch.state(["!selected"])  
    
    # 設定主題樣式
    apply_titlebar_theme()

    # 關閉視窗即終止
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

def open_log_window():
    '''
    open a new window to display logs.\n
    開啟新視窗顯示程式運作訊息\n
    使用多執行緒(Thread)方式載入後續動作, 避免tkinter無回應

    Args:
        NA.
    Return:
        NA.    
    '''    
    global log_text
    
    # tk視窗設定
    log_window = tk.Toplevel()
    log_window.title("104人力銀行爬蟲程式")
    window_width = log_window.winfo_screenwidth()    # 取得螢幕寬度
    window_height = log_window.winfo_screenheight()  # 取得螢幕高度
    width = 800
    height = 320
    left = int((window_width - width)/2)       # 計算左上 x 座標
    top = int((window_height - height)/2)      # 計算左上 y 座標
    log_window.geometry(f"{width}x{height}+{left}+{top}")

    # title
    tk.Label(log_window, text="⏳運行中...", font = ("DFKai-SB", 11)).pack(side=tk.TOP, anchor=tk.NW)
    text_frame = ttk.Frame(log_window)
    text_frame.pack(fill="both", expand=True)

    # log window
    scrollbar = ttk.Scrollbar(text_frame)
    scrollbar.pack(side="right", fill="y")

    # log text config
    log_text = tk.Text(text_frame, wrap="word", font=("DFKai-SB", 11), yscrollcommand=scrollbar.set)
    log_text.pack(fill="both", expand=True)
    scrollbar.config(command=log_text.yview)

    # thread載入後續動作, 避免tkinter無回應
    Thread(target=driver_setup, daemon=True).start()

    # 關閉視窗即終止
    log_window.protocol("WM_DELETE_WINDOW", on_closing)


def toggle_theme():
    '''
    toggles between dark and light themes and updates the switch button"s text.\n
    切換深淺主題並更新滑桿文字

    Args:
        NA.
    Return:
        NA.
    '''
    current_theme = sv_ttk.get_theme()
    
    if current_theme == "dark":
        sv_ttk.set_theme("light")
        # 主題切換為淺色，文字改為「淺色模式」
        theme_switch.configure(text="淺色模式")
    else:
        sv_ttk.set_theme("dark")
        # 主題切換為深色，文字改為「深色模式」
        theme_switch.configure(text="深色模式")

    # 設定button style
    style = ttk.Style()
    style.configure("TButton", font=("DFKai-SB", 20))

    apply_titlebar_theme()

def apply_titlebar_theme():
    '''
    applies the theme to the window's title bar on supported Windows versions.\n
    根據Windows版本設定視窗標題列樣式

    Args:
        NA.
    Return:
        NA.
    '''
    version = sys.getwindowsversion()
    if version.major == 10 and version.build >= 22000:
        change_header_color(root, "#1c1c1c" if sv_ttk.get_theme() == "dark" else "#fafafa")
    elif version.major == 10:
        apply_style(root, "dark" if sv_ttk.get_theme() == "dark" else "normal")
        root.wm_attributes("-alpha", 0.99)
        root.wm_attributes("-alpha", 1)

def on_start():
    '''
    when the "Start" button is clicked, do.
    當按下開始按鈕後檢查欄位是否為null\n
    若為null則跳出警告視窗\n
    若不為null則隱藏主視窗並開啟日誌視窗

    Args:
        NA.
    Return:
        NA.
    '''
    global keyword
    keyword = keyword_pack.get().strip()
    if not keyword:
        messagebox.showwarning("警告", "❗請輸入關鍵字❗")
        return
    root.withdraw()  # 隱藏主視窗
    open_log_window()  # 開啟日誌視窗

def on_closing():
    '''
    when the window is closed, and it terminates the entire program.\n
    當視窗關閉時, 終止整個程式

    Args:
        NA.
    Return:
        NA.
    '''
    root.destroy()
    sys.exit(0)

def driver_setup():
    '''
    setup selenium webdriver.\n
    selenium driver設定\n
    使用瀏覽器 - Chrome

    Args:
        keyword (str): 關鍵字
    Return:
        NA.    
    '''
    global driver
    # driver設定
    options = webdriver.ChromeOptions()
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_argument("-ignore-certificate-errors")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--user-agent='Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/12.0.3 Safari/605.1.15'")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])  # 屏蔽大部分無關 log
    # 爬蟲參數設定
    try:
        # 取得tk輸入的關鍵詞
        print_message(f"關鍵字為: {keyword}, 開啟瀏覽器中, 請勿操作瀏覽器, 會導致錯誤", "driver_setup", "info")
        url = f"https://www.104.com.tw/jobs/search/?jobsource=joblist_search&keyword={keyword}&mode=s&page=1&order=15&s9=1&isnew=30&searchJobs=1"
        driver = webdriver.Chrome(options = options)
        driver.get(url)
        # 休息10秒避免速度太快
        time.sleep(10)
        get_data()      
    except Exception as e:
        print_message(f"{e} ShutDown", "driver_setup", "err")
        exit(0)

def get_item_chinese(item):
    '''
    get item chinese.\n
    欄位中文轉換

    Args:
        item (str): 欄位名稱
    Return:
        str: 欄位中文名稱
    '''
    match item:
        case "title_cache":
            return "職缺名稱"
        case "companyname_cache":
            return "職缺公司名稱"
        case "url_cache":
            return "職缺網址資訊"
        case "companyindustry_cache":
            return "職缺公司所屬產業別"
        case "year_cache":
            return "年資要求"
        case "announcedate_cache":
            return "職缺公布時間"
        case "location_cache":
            return "工作地點"
        case "worktime_cache":
            return "上班時段"
        case "edu_cache":
            return "學歷要求"
        case "salary_cache":
            return "待遇資訊"
        case "content_cache":
            return "工作內容"
        case "category_cache":
            return "職務類別"
        case "dept_cache":
            return "科系要求"
        case "specialty_cache":
            return "擅長工具"
        case "others_cache":
            return "其他條件"            

def change_page(page):
    '''
    change page button to next page.\n
    尋找並點擊換頁按鈕

    Args:
        page (int): 頁數
    Return:
        NA.
    '''
    try:
        driver.get(f"https://www.104.com.tw/jobs/search/?jobsource=joblist_search&keyword={keyword}&mode=s&page={page}&order=15&s9=1&isnew=30&searchJobs=1")
        time.sleep(10)
    except Exception as e:
        print_message(f"{e} ShutDown", "change_page", "err")
        exit(0)

def listbox_quantity():
    '''
    get listbox quantity.\n
    取得搜尋結果頁面數量

    Args:
        NA.
    Return:
        int: 頁數
    '''
    try:
        # 找到下拉按鈕
        down_icon = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((
                By.XPATH, "//i[@class='jb_icon_down ml-1']"))
        )
        down_icon.click()
        # 找尋下拉選單中的選項
        listbox_cache = driver.find_element(
            By.XPATH, "//div[@class='high-light multiselect__content-wrapper']"
        )
        listbox_cache = driver.find_element(
            By.XPATH, "//ul[@class='multiselect__content w-100']"
        )
        listbox=listbox_cache.find_elements(By.XPATH, "//li[@class='multiselect__element position-relative']")
        
        return len(listbox)
    except NoSuchElementException:
        print_message("無法取得頁數 ShutDown", "listbox_quantity", "err")
        exit(0)

def click_search():
    '''
    click search button.\n
    點擊搜尋按鈕讓網頁載入正常結果

    Args:
        NA.
    Return:
        NA.
    '''
    retry = 0
    maxretries = 3
    while retry < maxretries:
        try:
            searchBTN = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable(
                (By.XPATH, "//button[@class='btn btn-secondary btn-block btn-lg' and text()='搜尋']"))
            )
            searchBTN.click()
            time.sleep(10)
            break
        except NoSuchElementException:
            if retry >= maxretries:
                raise Exception(f"搜尋按鈕未找到 重試次數: {retry + 1}")
            retry += 1
            time.sleep(2)
        except Exception as e:
            print_message(f"{e} ShutDown", "click_search", "err")
            exit(0)
 

def append_data1(condition, col, item):
    '''
    append data to the list.\n
    將搜尋結果列表中的部分資料放入list()

    Args:
        condition (tuple): element定位條件
        col (list): 欲放入資料之list
        item (str): 欄位名稱
    Return:
        NA.
    '''
    retry = 0
    max_retry = 3
    while retry < max_retry:
        try:
            for data in driver.find_elements(*condition):
                col.append(
                    data.text.replace("\n", " ")
                             .replace("學歷要求", "")
                             .replace("科系要求", "")
                             .replace("擅長工具", "")
                             .replace("提升專業能力", "")
                             .replace("其他條件", "")
                             .strip()
                )
            break
        except NoSuchElementException:
            retry += 1
            time.sleep(2)
            if retry >= max_retry:
                item_chinese1 = get_item_chinese(item)
                print_message(f"{e} element is {item_chinese1}, delete lost row", "append_data1", "err")
                deal_noelement(item)
                break
        except Exception as e:
            print_message(f"{e} ShutDown", "append_date1", "err")
            exit(0)

def deal_noelement(item):
    '''
    delete lost data.\n
    刪除因找不到element而導致資料長度不一致之資料

    Args:
        item (str): 欄位名稱
    Return:
        NA.
    '''
    try:
        match item:
            case "titles_cache":
                pass
            case "companyname_cache":
                job_titles.pop()
            case "url_cache":
                job_titles.pop()
                job_companyname.pop()
            case "companyindustry_cache":
                job_titles.pop()
                job_companyname.pop()
                job_url.pop()       
            case "year_cache":
                job_titles.pop()
                job_companyname.pop()
                job_url.pop()
                job_companyindustry.pop()
        print_message("deleted lost data, continue", "deal_noelement", "info")
    except Exception as e:
        print_message(f"{e} ShutDown", "deal_noelement", "err")
        exit(0)            


def show_list_status(page):
    '''
    show list status.\n
    顯示目前各list長度作為檢查

    Args:
        page (int): 頁數
    Return:
        NA.
    '''
    print(f"第{page}頁資料擷取完成")
    print(f"職缺名稱: {len(job_titles)}")          
    print(f"職缺公司名稱: {len(job_companyname)}")     
    print(f"職缺網址資訊: {len(job_url)}")             
    print(f"職缺公司所屬產業別: {len(job_companyindustry)}") 
    print(f"待遇資訊: {len(job_salary)}")                
    print(f"年資要求: {len(job_year)}")
    print(f"職缺公布時間: {len(job_announcedate)}") 
    print(f"工作地點: {len(job_location)}")
    print(f"上班時段: {len(job_worktime)}") 
    print(f"學歷要求: {len(job_edu)}")     
    print(f"工作內容: {len(job_content)}")         
    print(f"職務類別: {len(job_category)}")      
    print(f"科系要求: {len(job_dept)}")    
    print(f"擅長工具: {len(job_specialty)}")      
    print(f"其他條件: {len(job_others)}")
    print("--------------------------------------------------")

def append_data2(condition, col, item):
    '''
    append data to the list.\n
    將職缺詳細資料中的部分資料放入list()

    Args:
        condition (tuple): element定位條件
        col (list): 欲放入資料之list
        item (str): 欄位名稱
    Return:
        NA.
    '''
    retry = 0
    max_retry = 3
    while retry < max_retry:
        try:
            data = driver.find_element(*condition)
            col.append(
                data.text.replace("\n", "")
                        .replace("學歷要求", "")
                        .replace("科系要求", "")
                        .replace("擅長工具", "")
                        .replace("工作待遇", "")
                        .replace("提升專業能力", "")
                        .replace("/", "、")
                        .replace("其他條件", "")
                        .replace("（經常性薪資達 4 萬元或以上）", "")
                        .replace("取得專屬你的薪水報告", "")
                        .replace("（固定或變動薪資因個人資歷或績效而異）", "")
                        .strip()
            )
            break
        except NoSuchElementException:
            retry += 1
            time.sleep(2)
            if retry >= max_retry:
                col.append("")
                item_chinese2 = get_item_chinese(item)
                print_message(f"element is {item_chinese2}, append null value, continue", "append_data2", "err")
                break
                # raise Exception(f"over try to get detail data. append null value")
        except Exception as e:
            print_message(f"{e} ShutDown", "append_date2", "err")
            exit(0)      

def append_url(condition, col):
    '''
    append url to the list.\n
    將職缺網址放入list()

    Args:
        condition (tuple): element定位條件
        col (list): 欲放入資料之list
    Return:
        NA.
    '''
    retry = 0
    max_retry = 3
    while retry < max_retry:
        try:
            for data in driver.find_elements(*condition):
                col.append(
                    data.get_attribute("href")
                )
                url_cache.append(
                    data.get_attribute("href")
                )
            break
        except NoSuchElementException:
            if retry < max_retry:
                retry += 1
                time.sleep(2)
            else:
                print_message("over try to get url", "append_url", "err")
                col.append("")
                break
        except Exception as e:
            print_message(e, "append_url", "err")
            col.append("")
            break     

def append_date(condition, col):
    '''
    append date to the list.\n
    將職缺公布時間放入list()

    Args:
        condition (tuple): element定位條件
        col (list): 欲放入資料之list
    Return:
        NA.
    '''
    retry = 0
    max_retry = 3
    while retry < max_retry:
        try:
            data = driver.find_element(*condition)
            col.append(
                data.text.replace("更新", "").strip()
            )
            break
        except NoSuchElementException:
            if retry < max_retry:
                retry += 1
                time.sleep(2)
            else:
                print_message("over try to get date", "append_date", "err")
                col.append("")
                break
        except Exception as e:
            print_message(e, "append_date", "err")
            col.append("")
            break

def adj_excel(path):
    '''
    adjust excel format.\n
    調整excel欄寬

    Args:
        path (str): excel檔案路徑
    Return:
        NA.
    '''
    try:
        wb = load_workbook(path)
        ws = wb.active

        for column_cells in ws.columns:
            # 計算該欄最大長度
            max_length = 0
            column = column_cells[0].column_letter 
            for cell in column_cells:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            adjusted_width = max_length + 2  
            ws.column_dimensions[column].width = adjusted_width

        wb.save(path)
    except Exception as e:
        print_message(f"{e} ShutDown", "adj_excel", "err")
        exit(0)        

def concat_df(time):
    '''
    concat DataFrame.\n
    將各list合併成DataFrame並輸出至excel
    
    Args:
        time (str): 狀態參數
    Return:
        NA.
    '''
    try:
        df_lists = {"職缺名稱"          : job_titles,
                    "職缺名稱"          : job_companyname,
                    "職缺網址資訊"      : job_url,
                    "職缺公司所屬產業別" : job_companyindustry,
                    "職缺公布時間"      : job_announcedate,
                    "工作地點"          : job_location,
                    "年資要求"          : job_year,
                    "學歷要求"          : job_edu,
                    "待遇資訊"          : job_salary,
                    "工作內容"          : job_content,
                    "職務類別"          : job_category,
                    "科系要求"          : job_dept,
                    "擅長工具"          : job_specialty,
                    "其他條件"          : job_others

        }
        lengths = {k: len(v) for k, v in df_lists.items()}

        if len(set(lengths.values())) != 1:
            match time:
                case "0":
                    print_message(f"各欄位長度不一致: {lengths}, 重抓資料", "concat_df", "war")
                    # 重抓資料
                    job_titles.clear()
                    job_companyname.clear()
                    job_url.clear()
                    job_companyindustry.clear()
                    job_year.clear()
                    job_announcedate.clear()
                    job_location.clear()
                    job_worktime.clear()
                    job_edu.clear()
                    job_salary.clear()
                    job_content.clear()
                    job_category.clear()
                    job_dept.clear()
                    job_specialty.clear()
                    job_others.clear()
                    get_data()
                    concat_df("1")
                case "1":
                    print_message(f"各欄位長度不一致: {lengths}, ShutDown", "concat_df", "err")
                    exit(0)            
        else:
            now = datetime.now()
            data = pd.DataFrame(df_lists)
            data.to_excel(path.join(download_path, f"104人力銀行_{keyword}{now.strftime("%Y%m%d")}結果.xlsx"), index=False)
            driver.quit()
            adj_excel(path.join(download_path, f"104人力銀行_{keyword}{now.strftime("%Y%m%d")}結果.xlsx"))
            print_message(f"導出成功工作職缺資訊已儲存至: {download_path}", "concat_df", "info")
    except Exception as e:
        print_message(f"{e} ShutDown", "concat_df", "err")
        exit(0)
                
def get_data():
    '''
    get data from the page.\n
    取得104人力銀行職缺資訊\n
    分兩階段取得資料\n
    第一階段 - 取得搜尋結果列表中的部分資料\n
    第二階段 - 進入職缺詳細資料頁面取得剩餘資料
    
    Args:
        NA.
    Return:
        NA.
    '''
    print_message("開始取得104人力銀行職缺資訊...", "get_data", "info")           
    # 網頁直接進入可能會是空架構, 故再次點擊搜尋介面, 讓網頁載入正常結果
    cache_list1 = [
        "title_cache",
        "companyname_cache",
        "url_cache",
        "companyindustry_cache",
        "year_cache",
    ]
    cache_list2 = [
        "announcedate_cache",
        "location_cache",
        "worktime_cache",
        "edu_cache",
        "salary_cache",
        "content_cache",
        "category_cache",
        "dept_cache",
        "specialty_cache",
        "others_cache"
    ]
    try:
        click_search()
        # 取得搜尋結果頁面數量
        pages = listbox_quantity()
        print_message(f"搜尋結果共{pages}頁, 一頁抓取耗時約2分鐘, 抓取完成需{pages*2}分鐘", "get_data", "info")       
        for page in range(1, pages + 1): 
            
            print_message(f"正在取得第{page}頁資料中...", "get_data", "info")

            for item in cache_list1:

                match item:

                    case "title_cache":
                        condition = (By.XPATH, "//a[@data-gtm-joblist='職缺-職缺名稱']")
                        col = job_titles
                    case "companyname_cache":
                        condition = (By.XPATH, "//a[@data-gtm-joblist='職缺-公司名稱']")
                        col = job_companyname
                    case "url_cache":
                        condition = (By.XPATH, "//a[@data-gtm-joblist='職缺-職缺名稱']")
                        col = job_url
                        append_url(condition, col)
                        continue
                    case "companyindustry_cache":
                        condition = (By.XPATH, "//span[@class='info-company-addon-type text-gray-darker font-weight-bold']")
                        col = job_companyindustry
                    # case "salary_cache":
                    #     condition = (By.XPATH, "//a[contains(@data-gtm-joblist, '職缺-薪資')]")
                    #     col = job_salary
                    case "year_cache":
                        condition = (By.XPATH, "//a[contains(@data-gtm-joblist, '職缺-經歷')]")
                        col = job_year        
                append_data1(condition, col, item)                
            time.sleep(1)
            if url_cache == []:
                print(f"在page = {page}時 URL為空值")
            for url in url_cache:
                driver.get(url)
                time.sleep(3)    
                for item in cache_list2:               
                    match item:

                        case "announcedate_cache":
                            condition = (By.XPATH, "//span[@class='mx-3 t4 text-gray-darker' and contains(text(), '更新')]")
                            col = job_announcedate
                            append_date(condition, col)
                            continue
                        case "location_cache":
                            condition = (By.CSS_SELECTOR, "span[data-v-e81f764d]")
                            col = job_location
                        case "worktime_cache":
                            condition = (By.XPATH, "//div[contains(@class, 't3 mb-0') and contains(text(), '班')]")
                            col = job_worktime
                        case "edu_cache":
                            condition = (By.XPATH, "//h3[normalize-space(text())='學歷要求']/ancestor::div[contains(@class,'list-row')]")
                            col = job_edu
                        case "salary_cache":
                            condition = (By.XPATH, "//h3[normalize-space(text())='工作待遇']/ancestor::div[contains(@class,'list-row')]")
                            col = job_salary    
                        case "content_cache":
                            condition = (By.XPATH, "//p[@class='mb-5 r3 job-description__content text-break']")
                            col = job_content
                        case "category_cache":
                            condition = (By.XPATH, "//div[@class='category-item col p-0 job-description-table__data']")
                            col = job_category
                        case "dept_cache":
                            condition = (By.XPATH, "//h3[normalize-space(text())='科系要求']/ancestor::div[contains(@class,'list-row')]")
                            col = job_dept
                        case "specialty_cache":
                            condition = (By.XPATH, "//h3[normalize-space(text())='擅長工具']/ancestor::div[contains(@class,'list-row')]")
                            col = job_specialty
                        case "others_cache":
                            condition = (By.XPATH, "//h3[normalize-space(text())='其他條件']/ancestor::div[contains(@class,'list-row')]")
                            col = job_others        
                    append_data2(condition, col, item)   
                time.sleep(1)
            url_cache.clear()
            #----------
            show_list_status(page)
            #----------    
            if page < pages:   
                # 換頁
                change_page(page + 1)      
            print_message(f"第{page}頁資料擷取完成", "get_data", "info")    
        concat_df("0")

    except Exception as e:
        print_message(e, "get_data", "err")
        driver.quit()    

# 重試次數設定
maxretries = 3

# folder設定
user_folder = path.expanduser("~")
download_path = path.join(user_folder, "Downloads")  

# Start
open_main_window()