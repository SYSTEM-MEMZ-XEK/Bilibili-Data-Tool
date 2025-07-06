import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import re
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
import json
import time
import random
import datetime
import os
from io import BytesIO
import tempfile
import threading
import webbrowser
import traceback
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementNotInteractableException, SessionNotCreatedException

# 全局配置
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
    'Referer': 'https://www.bilibili.com/'
}

def format_duration(seconds):
    """将秒数格式化为时分秒格式"""
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    seconds = seconds % 60
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

class BilibiliDataTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Bilibili视频数据爬虫工具")
        self.root.geometry("900x650")
        self.root.resizable(True, True)
        
        # 初始化变量
        self.driver = None
        self.bvids = set()
        self.stop_scraping_flag = False
        self.stop_exporting_flag = False
        self.scraping_thread = None
        self.exporting_thread = None
        
        # 创建UI
        self.create_widgets()
        
        # 加载配置
        self.load_config()
        
        # 设置日志回调
        self.log_callback = None
    
    def create_widgets(self):
        """创建UI控件"""
        # 创建标签页
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建BVID爬取标签页
        self.scrape_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.scrape_frame, text="BVID爬取")
        self.create_scrape_tab()
        
        # 创建视频导出标签页
        self.export_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.export_frame, text="视频信息导出")
        self.create_export_tab()
        
        # 状态栏
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)
        self.update_status("就绪")
    
    def create_scrape_tab(self):
        """创建BVID爬取标签页"""
        # 目标用户UID
        uid_frame = ttk.Frame(self.scrape_frame)
        uid_frame.pack(fill=tk.X, pady=5, padx=10)
        
        ttk.Label(uid_frame, text="目标用户UID:").pack(side=tk.LEFT, padx=(0, 10))
        self.uid_entry = ttk.Entry(uid_frame, width=30)
        self.uid_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Cookie输入
        cookie_frame = ttk.LabelFrame(self.scrape_frame, text="自定义Cookie（可选）")
        cookie_frame.pack(fill=tk.BOTH, expand=True, pady=5, padx=10)
        
        self.cookie_text = scrolledtext.ScrolledText(cookie_frame, height=6)
        self.cookie_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 状态信息
        status_frame = ttk.LabelFrame(self.scrape_frame, text="爬取状态")
        status_frame.pack(fill=tk.BOTH, expand=True, pady=5, padx=10)
        
        self.status_text = scrolledtext.ScrolledText(status_frame, height=10)
        self.status_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.status_text.config(state=tk.DISABLED)
        
        # 按钮区域
        button_frame = ttk.Frame(self.scrape_frame)  # 修复这里：使用ttk.Frame而不是ttk.FFrame
        button_frame.pack(fill=tk.X, pady=10, padx=10)
        
        self.start_scrape_button = ttk.Button(button_frame, text="开始爬取", command=self.start_scraping)
        self.start_scrape_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_scrape_button = ttk.Button(button_frame, text="停止爬取", command=self.stop_scraping, state=tk.DISABLED)
        self.stop_scrape_button.pack(side=tk.LEFT, padx=5)
        
        self.save_scrape_button = ttk.Button(button_frame, text="保存BVID", command=self.save_bvids, state=tk.DISABLED)
        self.save_scrape_button.pack(side=tk.LEFT, padx=5)
        
        self.clear_scrape_button = ttk.Button(button_frame, text="清除日志", command=self.clear_scrape_log)
        self.clear_scrape_button.pack(side=tk.RIGHT, padx=5)
    
    def create_export_tab(self):
        """创建视频信息导出标签页"""
        # 输入文件选择
        input_frame = ttk.Frame(self.export_frame)
        input_frame.pack(fill=tk.X, pady=5, padx=10)
        
        ttk.Label(input_frame, text="BVID文件:").pack(side=tk.LEFT, padx=(0, 10))
        self.input_entry = ttk.Entry(input_frame)
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        self.browse_button = ttk.Button(input_frame, text="浏览...", width=10, command=self.browse_input_file)
        self.browse_button.pack(side=tk.RIGHT)
        
        # 输出文件名
        output_frame = ttk.Frame(self.export_frame)
        output_frame.pack(fill=tk.X, pady=5, padx=10)
        
        ttk.Label(output_frame, text="输出文件:").pack(side=tk.LEFT, padx=(0, 10))
        self.output_entry = ttk.Entry(output_frame)
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.output_entry.insert(0, "bilibili_videos.xlsx")
        
        # 选项设置
        options_frame = ttk.LabelFrame(self.export_frame, text="导出选项")
        options_frame.pack(fill=tk.X, pady=5, padx=10)
        
        self.include_cover_var = tk.BooleanVar(value=True)
        self.cover_check = ttk.Checkbutton(options_frame, text="包含视频封面", variable=self.include_cover_var)
        self.cover_check.pack(side=tk.LEFT, padx=10, pady=5)
        
        self.include_desc_var = tk.BooleanVar(value=True)
        self.desc_check = ttk.Checkbutton(options_frame, text="包含视频描述", variable=self.include_desc_var)
        self.desc_check.pack(side=tk.LEFT, padx=10, pady=5)
        
        self.include_tags_var = tk.BooleanVar(value=True)
        self.tags_check = ttk.Checkbutton(options_frame, text="包含视频标签", variable=self.include_tags_var)
        self.tags_check.pack(side=tk.LEFT, padx=10, pady=5)
        
        # 导出状态
        export_status_frame = ttk.LabelFrame(self.export_frame, text="导出状态")
        export_status_frame.pack(fill=tk.BOTH, expand=True, pady=5, padx=10)
        
        self.export_status_text = scrolledtext.ScrolledText(export_status_frame, height=10)
        self.export_status_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.export_status_text.config(state=tk.DISABLED)
        
        # 导出按钮区域
        export_button_frame = ttk.Frame(self.export_frame)
        export_button_frame.pack(fill=tk.X, pady=10, padx=10)
        
        self.start_export_button = ttk.Button(export_button_frame, text="开始导出", command=self.start_exporting)
        self.start_export_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_export_button = ttk.Button(export_button_frame, text="停止导出", command=self.stop_exporting, state=tk.DISABLED)
        self.stop_export_button.pack(side=tk.LEFT, padx=5)
        
        self.open_export_button = ttk.Button(export_button_frame, text="打开文件夹", command=self.open_output_folder, state=tk.DISABLED)
        self.open_export_button.pack(side=tk.LEFT, padx=5)
        
        self.clear_export_button = ttk.Button(export_button_frame, text="清除日志", command=self.clear_export_log)
        self.clear_export_button.pack(side=tk.RIGHT, padx=5)
    
    def log_message(self, message, target="scrape"):
        """在指定区域显示日志消息"""
        if target == "scrape":
            self.status_text.config(state=tk.NORMAL)
            self.status_text.insert(tk.END, message + "\n")
            self.status_text.see(tk.END)
            self.status_text.config(state=tk.DISABLED)
        else:  # export
            self.export_status_text.config(state=tk.NORMAL)
            self.export_status_text.insert(tk.END, message + "\n")
            self.export_status_text.see(tk.END)
            self.export_status_text.config(state=tk.DISABLED)
        
        self.update_status(message)
    
    def update_status(self, message):
        """更新状态栏"""
        self.status_var.set(message)
        self.root.update()
    
    def clear_scrape_log(self):
        """清除爬取日志"""
        self.status_text.config(state=tk.NORMAL)
        self.status_text.delete(1.0, tk.END)
        self.status_text.config(state=tk.DISABLED)
    
    def clear_export_log(self):
        """清除导出日志"""
        self.export_status_text.config(state=tk.NORMAL)
        self.export_status_text.delete(1.0, tk.END)
        self.export_status_text.config(state=tk.DISABLED)
    
    def load_config(self):
        """加载上次使用的配置"""
        try:
            if os.path.exists("bilibili_tool_config.json"):
                with open("bilibili_tool_config.json", "r", encoding="utf-8") as f:
                    config = json.load(f)
                    self.uid_entry.insert(0, config.get("uid", ""))
                    self.cookie_text.insert(tk.END, config.get("cookie", ""))
                    self.input_entry.insert(0, config.get("input_file", ""))
                    self.output_entry.delete(0, tk.END)
                    self.output_entry.insert(0, config.get("output_file", "bilibili_videos.xlsx"))
                    
                    # 加载选项设置
                    self.include_cover_var.set(config.get("include_cover", True))
                    self.include_desc_var.set(config.get("include_desc", True))
                    self.include_tags_var.set(config.get("include_tags", True))
        except:
            pass
    
    def save_config(self):
        """保存当前配置"""
        try:
            config = {
                "uid": self.uid_entry.get(),
                "cookie": self.cookie_text.get("1.0", tk.END).strip(),
                "input_file": self.input_entry.get(),
                "output_file": self.output_entry.get(),
                "include_cover": self.include_cover_var.get(),
                "include_desc": self.include_desc_var.get(),
                "include_tags": self.include_tags_var.get()
            }
            with open("bilibili_tool_config.json", "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except:
            pass
    
    # ============================== BVID爬取功能 ==============================
    def start_scraping(self):
        """开始爬取BVID"""
        uid = self.uid_entry.get().strip()
        if not uid:
            messagebox.showerror("错误", "请输入目标用户UID")
            return
        
        # 禁用按钮
        self.start_scrape_button.config(state=tk.DISABLED)
        self.stop_scrape_button.config(state=tk.NORMAL)
        self.save_scrape_button.config(state=tk.DISABLED)
        self.stop_scraping_flag = False
        
        # 保存当前配置
        self.save_config()
        
        # 在新线程中运行爬取
        self.scraping_thread = threading.Thread(target=self.run_scraping, args=(uid,))
        self.scraping_thread.daemon = True
        self.scraping_thread.start()
    
    def stop_scraping(self):
        """停止爬取BVID"""
        self.stop_scraping_flag = True
        self.log_message("正在停止爬取...", "scrape")
        self.stop_scrape_button.config(state=tk.DISABLED)
    
    def save_bvids(self):
        """保存BVID结果"""
        if not self.bvids:
            messagebox.showinfo("提示", "没有可保存的BVID")
            return
        
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
                initialfile=f"bvid_{self.uid_entry.get().strip()}.txt"
            )
            if file_path:
                with open(file_path, "w", encoding="utf-8") as f:
                    for bvid in sorted(self.bvids):
                        f.write(f"{bvid}\n")
                messagebox.showinfo("成功", f"成功保存 {len(self.bvids)} 个BVID到 {file_path}")
                self.input_entry.delete(0, tk.END)
                self.input_entry.insert(0, file_path)
        except Exception as e:
            messagebox.showerror("错误", f"保存文件失败: {str(e)}")
    
    def setup_driver(self):
        """设置浏览器驱动"""
        try:
            options = Options()
            
            # 创建唯一的用户数据目录
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S%f")
            random_str = ''.join(random.choices('abcdefghijklmnopqrstuvwxyz0123456789', k=6))
            user_data_dir = os.path.abspath(f"edge_profile_{timestamp}_{random_str}")
            os.makedirs(user_data_dir, exist_ok=True)
            options.add_argument(f'--user-data-dir={user_data_dir}')
            
            # 可选：启用无头模式（如果需要）
            # options.add_argument('--headless')
            
            options.add_argument('--disable-bink-features=AutomationControlled')
            options.add_argument('--ignore-certificate-errors')
            options.add_argument('--log-level=3')
            options.add_argument('--window-size=1200,800')
            options.add_argument('--start-maximized')
            options.add_experimental_option('excludeSwitches', ['enable-automation'])
            options.add_argument('--disable-gpu')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--no-sandbox')
            
            driver_path = r'C:\msedgedriver.exe'
            service = Service(executable_path=driver_path)
            self.driver = webdriver.Edge(service=service, options=options)
            self.log_message("Edge浏览器驱动已成功启动", "scrape")
            
            # 绕过webdriver检测
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            
            return True
        except SessionNotCreatedException as e:
            # 特定处理会话创建错误
            if "user data directory is already in use" in str(e):
                self.log_message("错误：用户数据目录已被占用，请稍后重试", "scrape")
                self.log_message("解决方案：请等待之前的浏览器实例完全关闭", "scrape")
            else:
                self.log_message(f"无法启动Edge浏览器: {str(e)}", "scrape")
            return False
        except Exception as e:
            error_msg = f"无法启动Edge浏览器: {str(e)}"
            self.log_message(error_msg, "scrape")
            
            # 记录详细错误信息
            error_details = traceback.format_exc()
            self.log_message(f"详细错误信息:\n{error_details}", "scrape")
            return False

    def set_cookies(self):
        """设置自定义Cookie"""
        cookie_str = self.cookie_text.get("1.0", tk.END).strip()
        if not cookie_str:
            self.log_message("未提供自定义Cookie，使用默认设置", "scrape")
            return
        
        try:
            # 清除所有现有cookies
            self.driver.delete_all_cookies()
            
            # 解析cookies字符串
            cookies = []
            for cookie in cookie_str.split(';'):
                cookie = cookie.strip()
                if not cookie:
                    continue
                parts = cookie.split('=', 1)
                if len(parts) < 2:
                    continue
                name = parts[0].strip()
                value = parts[1].strip()
                cookies.append({
                    'name': name,
                    'value': value,
                    'domain': '.bilibili.com',
                    'path': '/',
                    'secure': False,
                    'httpOnly': False,
                    'sameSite': 'Lax'
                })
            
            # 添加每个cookie
            for cookie in cookies:
                try:
                    self.driver.add_cookie(cookie)
                except Exception as e:
                    self.log_message(f"添加cookie {cookie['name']} 失败: {str(e)}", "scrape")
            
            self.log_message(f"已成功设置 {len(cookies)} 个自定义Cookie", "scrape")
            return True
        except Exception as e:
            self.log_message(f"设置Cookie失败: {str(e)}", "scrape")
            return False
    
    def navigate_to_page(self, url):
        """导航到目标页面"""
        self.driver.get(url)
        self.log_message(f"已访问目标页面: {url}", "scrape")
        time.sleep(2)  # 初始等待
    
    def click_next_page(self):
        """点击下一页按钮"""
        if self.stop_scraping_flag:
            return False
        
        try:
            # 尝试定位下一页按钮
            next_btn = WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '下一页')]"))
            )
            
            # 滚动到按钮位置
            self.driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", next_btn)
            
            # 点击按钮
            next_btn.click()
            self.log_message("已点击下一页按钮", "scrape")
            
            # 等待页面加载
            time.sleep(2)
            return True
        except (TimeoutException, ElementNotInteractableException):
            self.log_message("未找到下一页按钮或按钮不可点击", "scrape")
            return False
        except Exception as e:
            self.log_message(f"点击下一页按钮时出错: {str(e)}", "scrape")
            return False
    
    def extract_bvids(self):
        """从当前页面提取所有BVID号"""
        try:
            # 获取页面源码
            page_source = self.driver.page_source
            
            # 使用正则表达式提取所有BVID
            bvid_pattern = r'/video/(BV\w{10})'
            matches = re.findall(bvid_pattern, page_source)
            
            if matches:
                # 添加到集合中（自动去重）
                new_bvids = set(matches)
                self.bvids.update(new_bvids)
                self.log_message(f"当前页面提取到 {len(new_bvids)} 个新BVID，总计 {len(self.bvids)} 个", "scrape")
                return True
            else:
                self.log_message("未找到任何BVID", "scrape")
                return False
        except Exception as e:
            self.log_message(f"提取BVID时出错: {str(e)}", "scrape")
            return False
    
    def run_scraping(self, uid):
        """执行爬取流程"""
        try:
            # 设置浏览器驱动
            if not self.setup_driver():
                self.log_message("无法启动浏览器，爬取终止", "scrape")
                # 启用按钮
                self.start_scrape_button.config(state=tk.NORMAL)
                self.stop_scrape_button.config(state=tk.DISABLED)
                return
            
            # 导航到B站主页设置cookies
            self.driver.get("https://www.bilibili.com")
            time.sleep(1)
            
            # 设置自定义Cookie
            self.set_cookies()
            
            # 导航到目标页面
            target_url = f"https://space.bilibili.com/{uid}/video"
            self.navigate_to_page(target_url)
            
            # 先提取第一页的BVID
            self.extract_bvids()
            
            page_count = 1
            self.log_message(f"开始第{page_count}页", "scrape")
            
            while not self.stop_scraping_flag and page_count <= 50:
                success = self.click_next_page()
                if not success:
                    self.log_message("已到达最后一页", "scrape")
                    break
                
                page_count += 1
                self.log_message(f"成功翻到第{page_count}页", "scrape")
                
                # 提取当前页的BVID
                self.extract_bvids()
                
                # 随机延迟防止检测
                time.sleep(random.uniform(1.0, 2.5))
            
            self.log_message(f"共翻页 {page_count-1} 次，提取到 {len(self.bvids)} 个唯一BVID", "scrape")
            
        except Exception as e:
            error_msg = f"爬取过程中出错: {str(e)}"
            self.log_message(error_msg, "scrape")
            
            # 记录详细错误信息
            error_details = traceback.format_exc()
            self.log_message(f"详细错误信息:\n{error_details}", "极取")
        finally:
            try:
                # 尝试关闭浏览器
                if self.driver:
                    try:
                        self.driver.quit()
                    except:
                        pass
                    try:
                        self.driver.close()
                    except:
                        pass
                    self.driver = None
                    self.log_message("浏览器已关闭", "scrape")
            except Exception as e:
                self.log_message(f"关闭浏览器时出错: {str(e)}", "scrape")
            
            # 启用按钮
            self.start_scrape_button.config(state=tk.NORMAL)
            self.stop_scrape_button.config(state=tk.DISABLED)
            self.save_scrape_button.config(state=tk.NORMAL)
            
            self.log_message("爬取任务完成！", "scrape")
            
            # 如果爬取到了BVID，自动填充到导出标签页
            if self.bvids:
                try:
                    temp_file = os.path.join(tempfile.gettempdir(), f"temp_bvid_{uid}.txt")
                    with open(temp_file, "w", encoding="utf-8") as f:
                        for bvid in sorted(self.bvids):
                            f.write(f"{bvid}\n")
                    self.input_entry.delete(0, tk.END)
                    self.input_entry.insert(0, temp_file)
                    self.log_message(f"BVID列表已临时保存到 {temp_file}，可在导出标签页使用", "scrape")
                except:
                    pass
    
    # ============================== 视频信息导出功能 ==============================
    def browse_input_file(self):
        """浏览输入文件"""
        file_path = filedialog.askopenfilename(
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        if file_path:
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, file_path)
    
    def start_exporting(self):
        """开始导出视频信息"""
        input_file = self.input_entry.get().strip()
        output_file = self.output_entry.get().strip()
        
        if not input_file:
            messagebox.showerror("错误", "请选择BVID文件")
            return
        
        if not output_file:
            messagebox.showerror("错误", "请输入输出文件名")
            return
        
        # 禁用按钮
        self.start_export_button.config(state=tk.DISABLED)
        self.stop_export_button.config(state=tk.NORMAL)
        self.open_export_button.config(state=tk.DISABLED)
        self.stop_exporting_flag = False
        
        # 保存当前配置
        self.save_config()
        
        # 在新线程中运行导出
        self.exporting_thread = threading.Thread(target=self.run_exporting, args=(input_file, output_file))
        self.exporting_thread.daemon = True
        self.exporting_thread.start()
    
    def stop_exporting(self):
        """停止导出视频信息"""
        self.stop_exporting_flag = True
        self.log_message("正在停止导出...", "export")
        self.stop_export_button.config(state=tk.DISABLED)
        self.update_status("导出已停止")
    
    def open_output_folder(self):
        """打开输出文件所在文件夹"""
        output_file = self.output_entry.get().strip()
        if output_file and os.path.exists(output_file):
            folder_path = os.path.dirname(output_file)
            if os.path.exists(folder_path):
                webbrowser.open(folder_path)
    
    def extract_bvid(self, url_or_id):
        """从URL或ID中提取BVID"""
        # 如果是BV号直接返回
        if url_or_id.startswith("BV") and len(url_or_id) == 12:
            return url_or_id
        
        # 从URL中提取BV号
        bvid_match = re.search(r"(BV[0-9A-Za-z]{10})", url_or_id)
        if bvid_match:
            return bvid_match.group(1)
        
        # 从URL中提取av号
        av_match = re.search(r"av(\d+)", url_or_id)
        if av_match:
            return f"AV{av_match.group(1)}"
        
        # 尝试直接作为BV号处理
        if re.match(r"BV[0-9A-Za-z]{10}", url_or_id):
            return url_or_id
        
        return None
    
    def get_video_info(self, bvid):
        """通过B站API获取视频信息"""
        # 在请求前检查是否停止
        if self.stop_exporting_flag:
            return None
        
        api_url = ""

        if bvid.startswith("BV"):
            api_url = f"https://api.bilibili.com/x/web-interface/view?bvid={bvid}"
        elif bvid.startswith("AV"):
            aid = bvid[2:]
            api_url = f"https://api.bilibili.com/x/web-interface/view?aid={aid}"
        else:
            return None
        
        try:
            # 使用更短的超时时间，以便及时响应停止请求
            response = requests.get(api_url, headers=HEADERS, timeout=5)
            response.raise_for_status()
            data = response.json()
            
            if data.get("code") != 0:
                return None
                
            video_data = data["data"]
            
            # 确保标签正确获取
            tags = []
            try:
                # 使用独立API获取标签
                tag_api = f"https://api.bilibili.com/x/web-interface/view/detail/tag?bvid={video_data['bvid']}"
                tag_response = requests.get(tag_api, headers=HEADERS, timeout=3)
                tag_response.raise_for_status()
                tag_data = tag_response.json()
                if tag_data.get("code") == 0:
                    tags = [t["tag_name"] for t in tag_data["data"]]
            except requests.RequestException:
                # 忽略请求错误
                pass
            
            # 如果标签API失败，尝试从主API获取
            if not tags and 'tags' in video_data:
                tags = [t['tag_name'] for t in video_data['tags']]
            
            # 如果仍然没有标签，尝试从HTML解析
            if not tags:
                try:
                    # 在请求前检查是否停止
                    if self.stop_exporting_flag:
                        return None
                        
                    url = f"https://www.bilibili.com/video/{video_data['bvid']}"
                    response = requests.get(url, headers=HEADERS, timeout=5)
                    response.raise_for_status()
                    soup = BeautifulSoup(response.text, 'html.parser')
                    
                    # 尝试从HTML中提取标签
                    tag_elements = soup.select('.tag-link')
                    tags = [tag.text.strip() for tag in tag_elements]
                except requests.RequestException:
                    # 忽略HTML解析错误
                    pass
            
            # 创建结果字典
            result = {
                'title': video_data['title'],
                'bvid': video_data['bvid'],
                'author': video_data['owner']['name'],
                'author_id': video_data['owner']['mid'],
                'publish_date': time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(video_data['pubdate'])),
                'duration': format_duration(video_data['duration']),  # 格式化时长
                'duration_sec': video_data['duration'],  # 保留秒数用于计算
                'desc': video_data['desc'],
                'views': video_data['stat']['view'],
                'danmaku': video_data['stat']['danmaku'],
                'likes': video_data['stat']['like'],
                'coins': video_data['stat']['coin'],
                'favorites': video_data['stat']['favorite'],
                'shares': video_data['stat']['share'],
                'reply': video_data['stat']['reply'],  # 评论数
                'cover_url': video_data['pic'],  # 封面URL
                'tags': ','.join(tags)  # 标签
            }
            
            return result
        except requests.RequestException:
            return None
        except (KeyError, ValueError, TypeError) as e:
            self.log_message(f"解析视频信息错误: {str(e)}", "export")
            return None

    def run_exporting(self, input_file, output_file):
        """执行导出视频信息流程"""
        try:
            # 读取BVID文件
            try:
                with open(input_file, "r", encoding="utf-8") as f:
                    bvids = [line.strip() for line in f.readlines() if line.strip()]
            except Exception as e:
                self.log_message(f"读取文件失败: {str(e)}", "export")
                return
            
            if not bvids:
                self.log_message("BVID文件为空，没有可导出的视频", "export")
                return
            
            # 创建Excel工作簿
            wb = Workbook()
            ws = wb.active
            ws.title = "Bilibili视频数据"
            
            # 设置列宽
            ws.column_dimensions['A'].width = 50   # 标题
            ws.column_dimensions['B'].width = 15    # BVID
            ws.column_dimensions['C'].width = 15    # 作者
            ws.column_dimensions['D'].width = 15    # 作者ID
            ws.column_dimensions['E'].width = 20    # 发布时间
            ws.column_dimensions['F'].width = 15    # 时长
            ws.column_dimensions['G'].width = 12    # 播放量
            ws.column_dimensions['H'].width = 12    # 弹幕数
            ws.column_dimensions['I'].width = 12    # 点赞数
            ws.column_dimensions['J'].width = 12    # 投币数
            ws.column_dimensions['K'].width = 12    # 收藏数
            ws.column_dimensions['L'].width = 12    # 分享数
            ws.column_dimensions['M'].width = 12    # 评论数
            
            # 设置表头
            headers = [
                "标题", "BVID", "作者", "作者ID", "发布时间", "时长", 
                "播放量", "弹幕数", "点赞数", "投币数", "收藏数", "分享数", "评论数"
            ]
            
            # 根据选项添加额外表头
            if self.include_tags_var.get():
                headers.append("标签")
                ws.column_dimensions['N'].width = 40  # 标签列宽
            
            if self.include_desc_var.get():
                headers.append("描述")
                ws.column_dimensions['O'].width = 80  # 描述列宽
            
            if self.include_cover_var.get():
                headers.append("封面")
                ws.column_dimensions['P'].width = 20  # 封面列宽
            
            # 写入表头
            ws.append(headers)
            
            total = len(bvids)
            success_count = 0
            fail_count = 0
            
            self.log_message(f"开始导出 {total} 个视频的信息...", "export")
            
            # 遍历每个BVID
            for idx, bvid in enumerate(bvids, 1):
                if self.stop_exporting_flag:
                    self.log_message("导出已停止", "export")
                    break
                
                # 提取有效的BVID
                valid_bvid = self.extract_bvid(bvid)
                if not valid_bvid:
                    self.log_message(f"跳过无效的BVID/AVID: {bvid}", "export")
                    fail_count += 1
                    continue
                
                self.log_message(f"正在处理视频 [{idx}/{total}]: {valid_bvid}", "export")
                
                # 获取视频信息
                video_info = self.get_video_info(valid_bvid)
                
                if not video_info:
                    self.log_message(f"无法获取视频信息: {valid_bvid}", "export")
                    fail_count += 1
                    continue
                
                # 构建行数据
                row_data = [
                    video_info['title'],
                    video_info['bvid'],
                    video_info['author'],
                    video_info['author_id'],
                    video_info['publish_date'],
                    video_info['duration'],
                    video_info['views'],
                    video_info['danmaku'],
                    video_info['likes'],
                    video_info['coins'],
                    video_info['favorites'],
                    video_info['shares'],
                    video_info['reply']
                ]
                
                # 添加标签
                if self.include_tags_var.get():
                    row_data.append(video_info['tags'])
                
                # 添加描述
                if self.include_desc_var.get():
                    row_data.append(video_info['desc'])
                
                # 添加封面
                if self.include_cover_var.get():
                    try:
                        # 下载封面
                        response = requests.get(video_info['cover_url'], headers=HEADERS, timeout=10)
                        response.raise_for_status()
                        
                        # 将图片添加到Excel
                        img = XLImage(BytesIO(response.content))
                        img.width = 120
                        img.height = 80
                        
                        # 添加图片到单元格
                        column_letter = chr(65 + len(row_data))  # 获取当前列的字母表示
                        ws.add_image(img, f"{column_letter}{idx + 1}")
                        
                        # 设置行高以匹配图片
                        ws.row_dimensions[idx + 1].height = 60
                        
                        row_data.append("封面已添加")
                    except Exception as e:
                        self.log_message(f"下载或添加封面失败: {str(e)}", "export")
                        row_data.append("封面添加失败")
                
                # 写入行数据
                ws.append(row_data)
                success_count += 1
                
                # 更新状态
                progress = f"进度: {idx}/{total} | 成功: {success_count} | 失败: {fail_count}"
                self.log_message(progress, "export")
                self.update_status(progress)
                
                # 随机延迟防止请求过快
                time.sleep(random.uniform(0.5, 1.5))
            
            # 保存Excel文件
            try:
                wb.save(output_file)
                self.log_message(f"成功导出 {success_count} 个视频信息到 {output_file}", "export")
                self.open_export_button.config(state=tk.NORMAL)
            except Exception as e:
                self.log_message(f"保存Excel文件失败: {str(e)}", "export")
            
        except Exception as e:
            error_msg = f"导出过程中出错: {str(e)}"
            self.log_message(error_msg, "export")
            
            # 记录详细错误信息
            error_details = traceback.format_exc()
            self.log_message(f"详细错误信息:\n{error_details}", "export")
        finally:
            # 启用按钮
            self.start_export_button.config(state=tk.NORMAL)
            self.stop_export_button.config(state=tk.DISABLED)
            self.log_message("导出任务完成！", "export")

# 主程序入口
if __name__ == "__main__":
    root = tk.Tk()
    app = BilibiliDataTool(root)
    root.mainloop()