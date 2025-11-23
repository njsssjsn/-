import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, filedialog
import json
import os
import time
from datetime import datetime
from PIL import Image
import requests
from io import BytesIO
import threading
import re
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import webbrowser

# 核心依赖：DrissionPage 用于浏览器自动化抓取
from DrissionPage import ChromiumPage, ChromiumOptions

# --- 初始化配置与环境检查 ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

DATA_FILE = "temu_data_erp_v3.json" 

# **V7.0 修复: 权限和图片目录强制切换到用户文档 (解决 WinError 5)**
IMG_DIR = os.path.join(os.path.expanduser("~"), "Documents", "TemuToolImages")
try:
    if not os.path.exists(IMG_DIR):
        os.makedirs(IMG_DIR)
except Exception as e:
    # 理论上 V7.0 已经修复了权限问题，但仍需警告
    messagebox.showerror("启动失败", f"无法创建图片存储目录: {IMG_DIR}\n错误: {e}")
    exit() 

# Matplotlib 中文乱码修复
plt.rcParams['font.sans-serif'] = ['SimHei']  
plt.rcParams['axes.unicode_minus'] = False  

# --- 货币配置与价格提取函数 ---
CURRENCIES = {
    '$': 'USD', '€': 'EUR', '£': 'GBP', 'CA$': 'CAD', 'C$': 'CAD', 'AU$': 'AUD'
}
DEFAULT_CURRENCY_SYMBOL = '$' 

def extract_price_and_symbol(html_content, default_symbol=DEFAULT_CURRENCY_SYMBOL):
    """
    从HTML中提取价格和货币符号，支持多币种。
    """
    price_pattern = re.compile(
        r'([$€£]|CA\$|C\$|AU\$)\s*(\d+[\d,.]*)', 
        re.IGNORECASE | re.DOTALL
    )
    
    matches = price_pattern.findall(html_content)
    
    if matches:
        symbol, price_str = matches[0]
        try:
            price = float(price_str.replace(',', '').strip('.'))
            return symbol, price
        except ValueError:
            pass 

    return default_symbol, 0.00


class TemuERPApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Temu 竞品数据情报局 (V7.1 稳定修复版)")
        self.geometry("1400x800") 

        self.products = self.load_data()
        self.current_img_data = None 
        # V7.0 修复: 彻底解决 pyimage 引用丢失问题
        self.image_refs = [] 
        self.preview_image_ref = None
        self.update_lock = threading.Lock() # 用于批量更新的线程锁
        self.manual_img_path = None # 用于手动上传的图片路径

        # --- 布局结构 ---
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.setup_sidebar()
        
        # 创建各个页面容器
        self.frame_dashboard = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.frame_input = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.frame_list = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.frame_analysis = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")

        self.init_dashboard()
        self.init_input()
        self.init_list()
        self.init_analysis()

        self.show_frame("list")

    # --- 数据加载/保存 ---
    def load_data(self):
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except: 
                messagebox.showwarning("数据警告", "数据文件已损坏或为空，已创建一个新的空数据列表。")
                return []
        return []

    def save_data(self):
        with open(DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(self.products, f, ensure_ascii=False, indent=4)

    # --- 导航侧边栏 ---
    def setup_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        
        ctk.CTkLabel(self.sidebar, text="Temu Intelligence", font=ctk.CTkFont(size=22, weight="bold")).pack(pady=40)
        
        self.btn_nav_dash = ctk.CTkButton(self.sidebar, text="📊 数据看板", height=40, fg_color="transparent", border_width=1, command=lambda: self.show_frame("dashboard"))
        self.btn_nav_dash.pack(pady=10, padx=20, fill="x")

        self.btn_nav_list = ctk.CTkButton(self.sidebar, text="📋 商品列表 (核心)", height=40, fg_color="#2196F3", command=lambda: self.show_frame("list"))
        self.btn_nav_list.pack(pady=10, padx=20, fill="x")

        self.btn_nav_input = ctk.CTkButton(self.sidebar, text="➕ 新增采集", height=40, fg_color="transparent", border_width=1, command=lambda: self.show_frame("input"))
        self.btn_nav_input.pack(pady=10, padx=20, fill="x")
        
        # 货币选择框
        ctk.CTkLabel(self.sidebar, text="默认货币:", font=ctk.CTkFont(size=14)).pack(pady=(50, 5))
        self.currency_var = ctk.StringVar(value=DEFAULT_CURRENCY_SYMBOL)
        self.currency_option_menu = ctk.CTkOptionMenu(self.sidebar, 
                                                     values=list(CURRENCIES.keys()), 
                                                     variable=self.currency_var)
        self.currency_option_menu.pack(padx=20, fill="x")
        
        # 新增功能: 批量更新
        self.btn_bulk_update = ctk.CTkButton(self.sidebar, text="⚡ 批量更新 (All)", height=40, fg_color="#FF9800", command=self.start_bulk_update_thread)
        self.btn_bulk_update.pack(pady=(20, 10), padx=20, fill="x")
        
        # 新增功能: 清空数据
        ctk.CTkButton(self.sidebar, text="🧹 清空所有数据", height=40, fg_color="#D32F2F", command=self.clear_all_data).pack(pady=(10, 50), padx=20, fill="x")
        
        ctk.CTkButton(self.sidebar, text="📥 导出Excel", height=40, fg_color="#00897B", command=self.export_to_excel).pack(pady=10, padx=20, fill="x", side="bottom")

    def show_frame(self, name):
        # 隐藏所有
        self.frame_dashboard.grid_forget()
        self.frame_input.grid_forget()
        self.frame_list.grid_forget()
        self.frame_analysis.grid_forget()
        
        # 重置按钮样式
        self.btn_nav_list.configure(fg_color="transparent", border_width=1)
        self.btn_nav_dash.configure(fg_color="transparent", border_width=1)
        self.btn_nav_input.configure(fg_color="transparent", border_width=1)

        if name == "dashboard":
            self.refresh_dashboard()
            self.frame_dashboard.grid(row=0, column=1, sticky="nsew")
            self.btn_nav_dash.configure(fg_color="#2196F3", border_width=0)
        elif name == "list":
            self.refresh_list()
            self.frame_list.grid(row=0, column=1, sticky="nsew")
            self.btn_nav_list.configure(fg_color="#2196F3", border_width=0)
        elif name == "input":
            self.frame_input.grid(row=0, column=1, sticky="nsew")
            self.btn_nav_input.configure(fg_color="#2196F3", border_width=0)
        elif name == "analysis":
            self.frame_analysis.grid(row=0, column=1, sticky="nsew")

    def init_dashboard(self):
        self.dash_content = ctk.CTkFrame(self.frame_dashboard, fg_color="transparent")
        self.dash_content.pack(fill="both", expand=True, padx=40, pady=40)
        
    def refresh_dashboard(self):
        for w in self.dash_content.winfo_children(): w.destroy()
        
        total_sku = len(self.products)
        total_sales_growth = 0
        hot_items = 0
        
        today = datetime.now().strftime("%Y-%m-%d")
        
        for p in self.products:
            hist = p.get('history', [])
            if len(hist) >= 2:
                growth = hist[-1]['sales'] - hist[-2]['sales']
                if hist[-1]['date'] == today: 
                     total_sales_growth += growth
            
            velocity = 0
            if len(hist) >= 2:
                d1 = datetime.strptime(hist[0]['date'], "%Y-%m-%d")
                d2 = datetime.strptime(hist[-1]['date'], "%Y-%m-%d")
                days = (d2 - d1).days
                if days == 0: days = 1
                total_growth = hist[-1]['sales'] - hist[0]['sales']
                velocity = round(total_growth / days, 1)
                
            if velocity > 50: hot_items += 1

        def create_card(parent, title, value, color):
            frame = ctk.CTkFrame(parent, fg_color=color, height=150)
            frame.pack(side="left", fill="x", expand=True, padx=10)
            ctk.CTkLabel(frame, text=title, font=("Arial", 16)).pack(pady=(30, 10))
            ctk.CTkLabel(frame, text=str(value), font=("Arial", 32, "bold")).pack(pady=10)

        row1 = ctk.CTkFrame(self.dash_content, fg_color="transparent")
        row1.pack(fill="x", pady=20)
        
        create_card(row1, "监控商品总数", total_sku, "#1E88E5")
        create_card(row1, "今日总销量增长", f"+{total_sales_growth}", "#43A047")
        create_card(row1, "🔥 潜在爆款", hot_items, "#E53935")


    # --- 2. 核心列表页 (List) ---
    def init_list(self):
        header = ctk.CTkFrame(self.frame_list, height=60)
        header.pack(fill="x", padx=20, pady=20)
        ctk.CTkLabel(header, text="商品监控列表", font=("Arial", 20, "bold")).pack(side="left", padx=20)
        ctk.CTkButton(header, text="🔄 刷新页面", width=100, command=self.refresh_list).pack(side="right", padx=10)

        # 表头 (使用 grid 布局)
        titles = ctk.CTkFrame(self.frame_list, height=40, fg_color="#333333")
        titles.pack(fill="x", padx=20)
        
        # V7.0 修复: 彻底固定列宽和对齐方式
        col_texts = ["图片", "商品信息", "🔥 价格", "当前总销量", "较上次增长", "日均销速", "状态", "操作", "删除"]
        col_weights = [0, 1, 0, 0, 0, 0, 0, 0, 0] # 只有商品信息列允许拉伸
        col_min_widths = [60, 300, 100, 100, 100, 100, 80, 120, 70]
        
        for i, text in enumerate(col_texts):
            titles.grid_columnconfigure(i, weight=col_weights[i], minsize=col_min_widths[i])
            ctk.CTkLabel(titles, text=text, anchor="w").grid(row=0, column=i, padx=(15 if i==0 else 10, 10), sticky="w")


        self.scroll_list = ctk.CTkScrollableFrame(self.frame_list)
        self.scroll_list.pack(fill="both", expand=True, padx=20, pady=(0, 20))

    def refresh_list(self):
        for w in self.scroll_list.winfo_children(): w.destroy()
        # V7.0 修复: 每次刷新前清空图片引用列表
        self.image_refs = [] 
        
        # V7.0 修复: 重新配置滚动列表中的列权重
        col_weights = [0, 1, 0, 0, 0, 0, 0, 0, 0] 
        col_min_widths = [60, 300, 100, 100, 100, 100, 80, 120, 70]
        
        for i, weight in enumerate(col_weights):
            self.scroll_list.grid_columnconfigure(i, weight=weight, minsize=col_min_widths[i])

        for i, p in enumerate(self.products):
            self.create_product_row(p, row_index=i)

    def open_link(self, url):
        """点击打开链接"""
        webbrowser.open(url)

    def create_product_row(self, p, row_index):
        # 使用 Grid 布局替代 Pack/Side，彻底解决错位问题
        row = ctk.CTkFrame(self.scroll_list, height=80)
        # V7.0 修复: 使用 grid 替代 pack，确保行内对齐和父容器一致
        row.grid(row=row_index, column=0, columnspan=9, sticky="ew", pady=2) 
        
        # 配置行内部的列权重
        col_weights = [0, 1, 0, 0, 0, 0, 0, 0, 0] 
        # V7.1 修复: 移除 action_frame 的 grid() 中的 minsize，并在这里统一设置
        col_min_widths = [60, 300, 100, 100, 100, 100, 80, 120, 70] 
        
        for i, weight in enumerate(col_weights):
            row.grid_columnconfigure(i, weight=weight, minsize=col_min_widths[i])


        col = 0
        
        # 1. 图片
        try:
            if os.path.exists(p.get('img_path', '')):
                pil_img = Image.open(p['img_path'])
                c_img = ctk.CTkImage(pil_img, size=(50, 50))
                # 关键：将图片引用保存到 self.image_refs
                self.image_refs.append(c_img) 
                ctk.CTkLabel(row, image=c_img, text="", width=60).grid(row=0, column=col, rowspan=2, padx=(15, 10), sticky="w")
            else:
                ctk.CTkLabel(row, text="无图", width=60).grid(row=0, column=col, rowspan=2, padx=(15, 10), sticky="w")
        except Exception: 
             ctk.CTkLabel(row, text="Err", width=60).grid(row=0, column=col, rowspan=2, padx=(15, 10), sticky="w")
        col += 1
        
        # 2. 标题和采集日
        title_label = ctk.CTkLabel(row, 
                                     text=p['name'][:35]+"...", 
                                     font=("Arial", 14, "bold"), 
                                     anchor="w",
                                     text_color="#2196F3", 
                                     cursor="hand2",
                                     width=300) 
        title_label.grid(row=0, column=col, padx=10, sticky="w")
        title_label.bind("<Button-1>", lambda event, url=p['url']: self.open_link(url))
        
        ctk.CTkLabel(row, text=f"采集日: {p['history'][0]['date']}", font=("Arial", 10), text_color="gray", anchor="w", width=300).grid(row=1, column=col, padx=10, sticky="w")
        col += 1
        
        # 3. 价格
        price_symbol = p.get('price_symbol', DEFAULT_CURRENCY_SYMBOL) 
        current_price = p.get('current_price', '-')
        price_text = f"{price_symbol}{current_price:.2f}" if isinstance(current_price, (int, float)) else '-'
        ctk.CTkLabel(row, text=price_text, font=("Arial", 14, "bold"), text_color="#FFD700", width=100, anchor="w").grid(row=0, column=col, rowspan=2, padx=10, sticky="w")
        col += 1

        # 4-7. 数据计算和显示 (保持不变)
        hist = p['history']; current_sales = p['current_sales']; growth = 0; velocity = 0
        if len(hist) >= 2:
            growth = hist[-1]['sales'] - hist[-2]['sales']
            d1 = datetime.strptime(hist[0]['date'], "%Y-%m-%d"); d2 = datetime.strptime(hist[-1]['date'], "%Y-%m-%d")
            days = (d2 - d1).days; days = 1 if days == 0 else days
            total_growth = hist[-1]['sales'] - hist[0]['sales']
            velocity = round(total_growth / days, 1)

        # 5. 总销量
        ctk.CTkLabel(row, text=str(current_sales), font=("Arial", 14), width=100, anchor="w").grid(row=0, column=col, rowspan=2, padx=10, sticky="w")
        col += 1

        # 6. 增长 
        color = "gray"; growth_text = "-"
        if len(hist) >= 2:
            if growth > 0: color = "#00E676"; growth_text = f"🔺 +{growth}"
            elif growth == 0: growth_text = "0"
            else: color = "#FF5252"; growth_text = str(growth)
        
        ctk.CTkLabel(row, text=growth_text, text_color=color, font=("Arial", 14, "bold"), width=100, anchor="w").grid(row=0, column=col, rowspan=2, padx=10, sticky="w")
        col += 1

        # 7. 销速
        ctk.CTkLabel(row, text=f"{velocity}/天", width=100, anchor="w").grid(row=0, column=col, rowspan=2, padx=10, sticky="w")
        col += 1

        # 8. 状态标签
        status = "观察"; s_color = "gray"
        if velocity > 50: status = "🔥 爆款"; s_color = "#D50000"
        elif velocity > 10: status = "📈 潜力"; s_color = "#FF9800"
        
        ctk.CTkLabel(row, text=status, text_color=s_color, font=("Arial", 12, "bold"), width=80, anchor="w").grid(row=0, column=col, rowspan=2, padx=10, sticky="w")
        col += 1

        # 9. 操作按钮 (分析/更新)
        action_frame = ctk.CTkFrame(row, fg_color="transparent")
        # V7.1 修复: 移除 minsize 参数 (该参数只用于 grid_columnconfigure)
        action_frame.grid(row=0, column=col, rowspan=2, padx=5, sticky="w") 
        
        ctk.CTkButton(action_frame, text="📈 分析", width=55, height=25, fg_color="#673AB7", command=lambda x=p: self.open_analysis(x)).pack(side="left", padx=2, pady=2)
        ctk.CTkButton(action_frame, text="🔄 更新", width=55, height=25, fg_color="#2196F3", command=lambda x=p: self.update_single_item(x)).pack(side="left", padx=2, pady=2)
        col += 1
        
        # 10. 删除按钮
        ctk.CTkButton(row, 
                      text="🗑️ 删除", 
                      width=70, 
                      height=25, 
                      fg_color="#F44336", 
                      hover_color="#D32F2F", 
                      command=lambda x=p: self.delete_product(x)).grid(row=0, column=col, rowspan=2, padx=10, sticky="w")

    # --- 新增功能: 批量更新 ---
    def start_bulk_update_thread(self):
        confirm = messagebox.askyesno("批量更新确认", f"您确定要更新列表中所有 {len(self.products)} 个商品吗？这可能需要几分钟。")
        if not confirm:
            return
        
        self.btn_bulk_update.configure(text="⚡ 批量更新中...", state="disabled")
        self.lbl_status.configure(text="开始批量更新...") # 使用输入页面的状态栏
        threading.Thread(target=self.bulk_update_logic, daemon=True).start()

    def bulk_update_logic(self):
        
        for i, p in enumerate(self.products):
            self.lbl_status.configure(text=f"批量更新: ({i+1}/{len(self.products)}) 正在处理 {p['name'][:20]}...")
            self.update_logic(p, is_bulk=True)
            # 休息一秒，防止爬取太快被封
            time.sleep(1) 
        
        # 完成后保存数据并刷新列表
        self.save_data()
        self.refresh_list()
        
        # 恢复状态
        self.btn_bulk_update.configure(text="⚡ 批量更新 (All)", state="normal")
        self.lbl_status.configure(text="批量更新完成！")
        messagebox.showinfo("批量更新", "所有商品数据已更新完毕。")

    # --- 核心更新逻辑 (用于单个和批量) ---
    def update_single_item(self, p):
        # 确保状态栏显示在输入页面
        self.show_frame("input")
        self.lbl_status.configure(text=f"正在更新: {p['name']}...")
        threading.Thread(target=self.update_logic, args=(p, False), daemon=True).start()

    def update_logic(self, p, is_bulk=False):
        try:
            co = ChromiumOptions().set_local_port(9222)
            page = ChromiumPage(co)
            page.get(p['url'])
            time.sleep(3)
            
            sales = 0; html = page.html
            patterns = [r'(\d+[\d,.]*)[+]*\s*sold', r'已售\s*(\d+[\d,.]*)[+]*']
            for pat in patterns:
                m = re.search(pat, html, re.IGNORECASE)
                if m:
                    num_str = m.group(1).replace(',', '')
                    if 'k' in num_str.lower(): sales = int(float(num_str.lower().replace('k',''))*1000)
                    else: sales = int(float(num_str))
                    break
            
            price_symbol, price_num = extract_price_and_symbol(page.html, p.get('price_symbol', self.currency_var.get()))
            price = price_num
            
            if sales > 0:
                with self.update_lock: # 确保在批量更新时数据写入是安全的
                    p['current_sales'] = sales; p['current_price'] = price 
                    p['price_symbol'] = price_symbol
                    
                    today = datetime.now().strftime("%Y-%m-%d")
                    new_history_data = {"date": today, "sales": sales, "price": price, "symbol": price_symbol}

                    if p['history'][-1]['date'] == today: p['history'][-1] = new_history_data
                    else: p['history'].append(new_history_data)
                    
                    if not is_bulk: # 单个更新时才立即保存和刷新
                        self.save_data()
                        self.refresh_list()
                        self.lbl_status.configure(text=f"更新完成: {p['name']}")
            
        except Exception as e:
            if not is_bulk:
                 self.lbl_status.configure(text=f"更新失败: {e}")
            else:
                 print(f"批量更新中，{p['name']} 失败: {e}") # 批量更新只打印错误不弹窗

    # --- 删除商品逻辑 ---
    def delete_product(self, product):
        confirm = messagebox.askyesno(
            "确认删除", 
            f"您确定要永久删除竞品: {product['name'][:30]}...? \n\n该操作无法撤销，相关图片也将被删除。"
        )
        
        if confirm:
            try:
                if os.path.exists(product.get('img_path', '')):
                    os.remove(product['img_path'])
            except Exception as e:
                print(f"删除图片失败: {e}")

            self.products = [p for p in self.products if p['id'] != product['id']]
            self.save_data()
            self.refresh_list()
            messagebox.showinfo("成功", "商品已删除。")
            
    # --- 新增功能: 清空所有数据 ---
    def clear_all_data(self):
        confirm = messagebox.askyesno(
            "⚠️ 严重警告：清空所有数据",
            f"您确定要删除所有 {len(self.products)} 个监控商品及其所有历史数据和图片吗？\n\n此操作**无法撤销**！"
        )
        if not confirm:
            return

        # 1. 删除所有图片文件
        deleted_count = 0
        for p in self.products:
            try:
                if os.path.exists(p.get('img_path', '')):
                    os.remove(p['img_path'])
                    deleted_count += 1
            except Exception as e:
                print(f"清空数据时，删除图片 {p.get('img_path')} 失败: {e}")

        # 2. 清空数据列表
        self.products = []
        
        # 3. 清空数据文件
        self.save_data()
        
        # 4. 刷新列表
        self.refresh_list()
        messagebox.showinfo("清空成功", f"所有数据已清空。共清理了 {deleted_count} 个图片文件。")


    # --- 3. 数据分析页 (Analysis) ---
    def init_analysis(self):
        self.ana_top = ctk.CTkFrame(self.frame_analysis, height=100)
        self.ana_top.pack(fill="x", padx=20, pady=20)
        self.ana_title = ctk.CTkLabel(self.ana_top, text="商品分析", font=("Arial", 20, "bold"))
        self.ana_title.pack(pady=20)
        
        ctk.CTkButton(self.ana_top, text="🔙 返回列表", command=lambda: self.show_frame("list")).place(x=20, y=30)
        
        self.chart_frame = ctk.CTkFrame(self.frame_analysis)
        self.chart_frame.pack(fill="both", expand=True, padx=20, pady=20)

    def open_analysis(self, product):
        self.show_frame("analysis")
        
        price_sym = product.get('price_symbol', DEFAULT_CURRENCY_SYMBOL)
        self.ana_title.configure(text=f"分析: {product['name']} (当前价格: {price_sym}{product.get('current_price', '-'):.2f})")
        
        for w in self.chart_frame.winfo_children(): w.destroy()
        
        hist = product['history']
        dates = [h['date'][5:] for h in hist] 
        sales = [h['sales'] for h in hist]
        prices = [h.get('price', None) for h in hist]
        
        if not dates:
            ctk.CTkLabel(self.chart_frame, text="暂无历史数据").pack()
            return

        fig, ax1 = plt.subplots(figsize=(8, 5), dpi=100, facecolor="#2b2b2b")
        fig.patch.set_facecolor('#2b2b2b') 
        ax1.set_facecolor("#2b2b2b")
        
        ax1.plot(dates, sales, marker='o', color='#00E676', linewidth=2, label='累计销量')
        ax1.set_xlabel("日期", color="white")
        ax1.set_ylabel("累计销量", color='#00E676')
        ax1.tick_params(axis='y', colors='#00E676')
        ax1.tick_params(axis='x', colors='white')
        
        plt.xticks(rotation=45, ha='right')

        if any(p is not None and p > 0 for p in prices):
            ax2 = ax1.twinx()
            ax2.plot(dates, prices, marker='x', linestyle='--', color='#FFD700', linewidth=1, label='商品价格')
            ax2.set_ylabel(f"价格 ({price_sym})", color='#FFD700')
            ax2.tick_params(axis='y', colors='#FFD700')
            ax2.spines['right'].set_color('#FFD700')
            ax2.spines['left'].set_color('#00E676')
        
        ax1.set_title("销量与价格趋势", color="white")
        plt.tight_layout()
        
        canvas = FigureCanvasTkAgg(fig, master=self.chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    # --- 4. 采集录入 (Input) ---
    def init_input(self):
        ctk.CTkLabel(self.frame_input, text="新增商品监控", font=("Arial", 20)).pack(pady=20)
        
        # =========================================================
        # V7.1 新增：手动录入功能区域
        # =========================================================
        manual_frame = ctk.CTkFrame(self.frame_input, fg_color="#333333", border_width=1, border_color="#FF9800")
        manual_frame.pack(fill="x", padx=50, pady=(0, 30))
        
        ctk.CTkLabel(manual_frame, text="✍️ 手动录入 (不依赖抓取)", font=("Arial", 16, "bold")).pack(pady=(15, 10))

        manual_form_frame = ctk.CTkFrame(manual_frame, fg_color="transparent")
        manual_form_frame.pack(fill="x", padx=20, pady=10)
        manual_form_frame.grid_columnconfigure(0, weight=1)
        manual_form_frame.grid_columnconfigure(1, weight=1)

        # 商品名称
        ctk.CTkLabel(manual_form_frame, text="商品名称:", anchor="w").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.manual_entry_title = ctk.CTkEntry(manual_form_frame, placeholder_text="必填：商品标题")
        self.manual_entry_title.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

        # 当前销量
        ctk.CTkLabel(manual_form_frame, text="当前销量:", anchor="w").grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.manual_entry_sales = ctk.CTkEntry(manual_form_frame, placeholder_text="必填：整数销量（例: 1250）")
        self.manual_entry_sales.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        
        # 价格
        ctk.CTkLabel(manual_form_frame, text="价格:", anchor="w").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        price_frame = ctk.CTkFrame(manual_form_frame, fg_color="transparent")
        price_frame.grid(row=3, column=0, padx=5, pady=5, sticky="ew")
        self.manual_entry_price_symbol = ctk.CTkEntry(price_frame, width=50, placeholder_text="$")
        self.manual_entry_price_symbol.pack(side="left", padx=(0, 5))
        self.manual_entry_price = ctk.CTkEntry(price_frame, placeholder_text="必填：价格（例: 12.99）", text_color="#FFD700") 
        self.manual_entry_price.pack(side="left", fill="x", expand=True)

        # 图片上传
        ctk.CTkLabel(manual_form_frame, text="图片上传 (可选):", anchor="w").grid(row=2, column=1, padx=5, pady=5, sticky="w")
        self.manual_btn_img = ctk.CTkButton(manual_form_frame, text="📂 选择图片 (本地)", command=self.select_manual_image)
        self.manual_btn_img.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        
        # 链接
        ctk.CTkLabel(manual_form_frame, text="商品链接 (可选):", anchor="w").grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        self.manual_entry_url = ctk.CTkEntry(manual_form_frame, placeholder_text="可选：商品链接")
        self.manual_entry_url.grid(row=5, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        # 保存按钮
        ctk.CTkButton(manual_frame, text="💾 手动入库", fg_color="#FF9800", hover_color="#E68A00", command=self.save_manual_product).pack(pady=15, padx=20, fill="x")

        ctk.CTkLabel(self.frame_input, text="--- 或使用下方自动抓取功能 ---").pack(pady=20)
        # =========================================================
        # V7.1 新增：手动录入功能区域结束
        # =========================================================
        
        input_box = ctk.CTkFrame(self.frame_input)
        input_box.pack(pady=20)
        
        self.entry_url = ctk.CTkEntry(input_box, width=500, placeholder_text="粘贴链接...")
        self.entry_url.pack(side="left", padx=10)
        
        self.btn_scrape = ctk.CTkButton(input_box, text="🚀 开始抓取", command=self.start_scrape_thread)
        self.btn_scrape.pack(side="left")
        
        self.lbl_status = ctk.CTkLabel(self.frame_input, text="等待操作...", text_color="gray")
        self.lbl_status.pack()
        
        # 预览结果区 (抓取结果)
        self.res_frame = ctk.CTkFrame(self.frame_input, height=300)
        self.res_frame.pack(fill="x", padx=50, pady=20)
        
        self.lbl_img_preview = ctk.CTkLabel(self.res_frame, text="图片预览")
        self.lbl_img_preview.pack(side="left", padx=20)
        
        form = ctk.CTkFrame(self.res_frame, fg_color="transparent")
        form.pack(side="left", fill="both", expand=True, padx=20)
        
        self.entry_title = ctk.CTkEntry(form, placeholder_text="标题")
        self.entry_title.pack(fill="x", pady=10)
        self.entry_sales = ctk.CTkEntry(form, placeholder_text="销量")
        self.entry_sales.pack(fill="x", pady=10)
        
        price_frame = ctk.CTkFrame(form, fg_color="transparent")
        price_frame.pack(fill="x", pady=10)
        self.entry_price_symbol = ctk.CTkEntry(price_frame, width=50, placeholder_text="$")
        self.entry_price_symbol.pack(side="left", padx=(0, 5))
        self.entry_price = ctk.CTkEntry(price_frame, placeholder_text="价格数字", text_color="#FFD700") 
        self.entry_price.pack(side="left", fill="x", expand=True)
        
        ctk.CTkButton(form, text="💾 确认入库", fg_color="green", command=self.save_product).pack(pady=20)

    # --- V7.1 新增: 手动录入辅助方法 ---
    def select_manual_image(self):
        """打开文件对话框选择本地图片并预览"""
        path = filedialog.askopenfilename(
            title="选择商品图片",
            filetypes=(("Image files", "*.png;*.jpg;*.jpeg"), ("All files", "*.*"))
        )
        if path:
            self.manual_img_path = path
            self.manual_btn_img.configure(text=os.path.basename(path))
            # 预览图片 (复用抓取页面的预览标签)
            try:
                pil = Image.open(path)
                c_img = ctk.CTkImage(pil, size=(150, 150))
                self.lbl_img_preview.configure(image=c_img, text="")
                self.preview_image_ref = c_img # 保存引用
            except Exception as e:
                messagebox.showerror("图片错误", f"无法加载图片: {e}")
                self.manual_img_path = None
                self.manual_btn_img.configure(text="📂 选择图片 (本地)")
        
    def save_manual_product(self):
        """手动录入并保存产品数据"""
        title = self.manual_entry_title.get()
        sales_str = self.manual_entry_sales.get()
        price_symbol = self.manual_entry_price_symbol.get() or self.currency_var.get()
        price_str = self.manual_entry_price.get()
        url = self.manual_entry_url.get() or "N/A (手动录入)"
        
        if not title or not sales_str or not price_str:
            messagebox.showerror("错误", "商品名称、销量和价格为必填项！"); return

        try:
            sales = int(sales_str); price = float(price_str)
            if sales < 0 or price < 0: raise ValueError
        except ValueError:
            messagebox.showerror("错误", "销量必须是整数，价格必须是有效数字！"); return
            
        pid = int(time.time()); path = ""

        # 处理图片
        if self.manual_img_path and os.path.exists(self.manual_img_path):
            try:
                # 复制图片到目标目录
                img_ext = os.path.splitext(self.manual_img_path)[1]
                path = os.path.join(IMG_DIR, f"{pid}{img_ext}")
                with Image.open(self.manual_img_path) as img:
                    img.save(path)
            except Exception as e:
                 messagebox.showerror("文件保存失败", f"无法保存图片到 {path}。请检查权限或路径。"); return
        
        new_p = {
            "id": pid, "name": title, "url": url, "img_path": path, 
            "current_sales": sales, "current_price": price, 
            "price_symbol": price_symbol, 
            "history": [{"date": datetime.now().strftime("%Y-%m-%d"), "sales": sales, "price": price, "symbol": price_symbol}] 
        }
        self.products.append(new_p); self.save_data()
        
        messagebox.showinfo("成功", "手动录入商品已入库！")
        
        # 清空手动录入的输入框和状态
        self.manual_entry_title.delete(0, "end"); self.manual_entry_sales.delete(0, "end")
        self.manual_entry_price.delete(0, "end"); self.manual_entry_price_symbol.delete(0, "end")
        self.manual_entry_url.delete(0, "end")
        self.manual_btn_img.configure(text="📂 选择图片 (本地)")
        self.manual_img_path = None
        self.lbl_img_preview.configure(image=None, text="图片预览") # 清空预览
        self.show_frame("list") # 自动跳转回列表页查看结果


    # 采集核心逻辑
    def start_scrape_thread(self):
        url = self.entry_url.get()
        if not url: return
        self.lbl_status.configure(text="正在连接浏览器...")
        threading.Thread(target=self.scrape_logic, args=(url,), daemon=True).start()

    def scrape_logic(self, url):
        try:
            co = ChromiumOptions().set_local_port(9222)
            page = ChromiumPage(co)
            page.get(url)
            self.lbl_status.configure(text="正在分析...")
            time.sleep(3)
            
            title = "未获取"
            try: title = page.ele('tag:h1').text
            except: pass
            
            sales = 0
            try:
                html = page.html
                patterns = [r'(\d+[\d,.]*)[+]*\s*sold', r'已售\s*(\d+[\d,.]*)[+]*']
                for p in patterns:
                    match = re.search(p, html, re.IGNORECASE)
                    if match:
                        num_str = match.group(1).replace(',', '')
                        if 'k' in num_str.lower(): sales = int(float(num_str.lower().replace('k',''))*1000)
                        else: sales = int(float(num_str))
                        break
            except: pass

            price_symbol, price_num = extract_price_and_symbol(page.html, self.currency_var.get())
            
            img_url = ""
            try:
                imgs = page.eles('tag:img')
                valid_imgs = [img for img in imgs if img.attr('src') and img.rect.size[0] > 300]
                if valid_imgs:
                     img_url = valid_imgs[0].attr('src')
            except: pass

            self.entry_title.delete(0, "end"); self.entry_title.insert(0, title)
            self.entry_sales.delete(0, "end"); self.entry_sales.insert(0, str(sales))
            self.entry_price_symbol.delete(0, "end"); self.entry_price_symbol.insert(0, price_symbol)
            self.entry_price.delete(0, "end"); self.entry_price.insert(0, f"{price_num:.2f}") 

            if img_url:
                if not img_url.startswith('http'): img_url = "https:"+img_url
                self.current_img_data = BytesIO(requests.get(img_url).content)
                pil = Image.open(self.current_img_data)
                c_img = ctk.CTkImage(pil, size=(150, 150))
                self.lbl_img_preview.configure(image=c_img, text="")
                self.preview_image_ref = c_img 

            self.lbl_status.configure(text="抓取成功！请确认后保存。")
            
        except Exception as e:
            self.lbl_status.configure(text=f"抓取失败，请检查链接或网络: {e}")

    def save_product(self):
        title = self.entry_title.get(); sales_str = self.entry_sales.get()
        price_symbol = self.entry_price_symbol.get(); price_str = self.entry_price.get()
        url = self.entry_url.get()

        try:
            sales = int(sales_str); price = float(price_str)
        except ValueError:
            messagebox.showerror("错误", "销量和价格必须是数字！"); return
        
        pid = int(time.time()); path = os.path.join(IMG_DIR, f"{pid}.jpg")
        if self.current_img_data:
            try:
                with open(path, "wb") as f: f.write(self.current_img_data.getbuffer())
            except Exception as e:
                messagebox.showerror("文件保存失败", f"无法保存图片到 {path}。请检查权限或路径。"); return

        new_p = {
            "id": pid, "name": title, "url": url, "img_path": path, 
            "current_sales": sales, "current_price": price, 
            "price_symbol": price_symbol, 
            "history": [{"date": datetime.now().strftime("%Y-%m-%d"), "sales": sales, "price": price, "symbol": price_symbol}] 
        }
        self.products.append(new_p); self.save_data()
        messagebox.showinfo("成功", "商品已入库！")
        
        self.entry_url.delete(0, "end"); self.entry_title.delete(0, "end")
        self.entry_sales.delete(0, "end"); self.entry_price_symbol.delete(0, "end")
        self.entry_price.delete(0, "end")
        self.lbl_img_preview.configure(image=None, text="图片预览")
        self.show_frame("list") # 采集成功后自动跳转回列表页查看结果

    # --- 6. 导出 Excel ---
    def export_to_excel(self):
        data = []
        for p in self.products:
            hist = p['history']
            growth = 0
            if len(hist) >= 2: growth = hist[-1]['sales'] - hist[-2]['sales']
            
            data.append({
                "商品名称": p['name'],
                "当前价格": f"{p.get('price_symbol', '')}{p.get('current_price', 0.0):.2f}",
                "当前销量": p['current_sales'],
                "较上次增长": growth,
                "采集链接": p['url'],
                "最后更新": hist[-1]['date']
            })
            
        df = pd.DataFrame(data)
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if path:
            df.to_excel(path, index=False)
            messagebox.showinfo("导出成功", f"文件已保存到: {path}")

if __name__ == "__main__":
    app = TemuERPApp()
    app.mainloop()
