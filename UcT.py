import sys
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import re
import time
import asyncio
import aiohttp
import concurrent.futures
from urllib.parse import urlparse
from collections import defaultdict
import zipfile
import io
import math
from functools import lru_cache
import unicodedata
import string
import base64
import tempfile
import traceback
import threading
import webbrowser
import pandas as pd
from datetime import datetime


# UcT核心配置
class UctConfig:
    MAX_WORKERS = 24  # 最大线程数(CPU核心数*3)
    MAX_ASYNC_TASKS = 256  # 异步任务最大并发数
    REQUEST_TIMEOUT = 10  # 请求超时时间(秒)
    USER_AGENT = "UcT/1.0 (Professional URL Checker)"  # 定制UA
    URL_REGEX = r'(https?://[^<\s\'"]{8,})'  # 高效URL识别正则
    CACHE_SIZE = 1024  # URL处理缓存大小
    CHUNK_SIZE = 50  # 批量结果处理块大小（提高效率）
    REQUIRED_PACKAGES = [
        'aiohttp',
        'pandas',
        'PyMuPDF',  # PDF处理库
        'openpyxl',
        'python-docx',
        'python-pptx'
    ]


# 静默依赖管理器
class SilentDependencyManager:
    @staticmethod
    def check_and_install_dependencies():
        """静默检查并安装缺失的依赖库"""
        try:
            # 获取已安装的包
            installed = {}
            try:
                import pkg_resources
                installed = {pkg.key for pkg in pkg_resources.working_set}
            except:
                # 如果pkg_resources不可用，使用其他方法
                try:
                    from importlib.metadata import distributions
                    installed = {dist.metadata['Name'].lower() for dist in distributions()}
                except:
                    return False

            # 检查哪些包缺失
            missing = []
            for package in UctConfig.REQUIRED_PACKAGES:
                normalized_name = package.replace('-', '_').lower()
                if normalized_name not in installed:
                    missing.append(package)

            if not missing:
                return True

            # 尝试自动安装
            print(f"静默安装缺失依赖: {', '.join(missing)}")
            python = sys.executable
            subprocess.check_call(
                [python, '-m', 'pip', 'install', *missing],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
            return True
        except Exception as e:
            print(f"静默安装失败: {str(e)}")
            return False


# UcT文件处理引擎
class UctEngine:
    @staticmethod
    @lru_cache(maxsize=UctConfig.CACHE_SIZE)
    def normalize_url(url):
        """高效URL规范化处理"""
        try:
            # URL预处理
            url = url.strip().rstrip('.,:;!?')
            url = unicodedata.normalize('NFC', url)

            # 添加协议（如果需要）
            if not url.startswith(('http://', 'https://')):
                if url.startswith('//'):
                    url = 'https:' + url
                elif '://' not in url and '.' in url:
                    url = 'https://' + url
                else:
                    return None

            # URL解析和重构
            parsed = urlparse(url)
            if not parsed.netloc:
                return None

            # 标准化域名（小写）
            netloc = parsed.netloc.lower()

            # 重建无参数URL
            clean_url = f"{parsed.scheme}://{netloc}{parsed.path}"
            return clean_url
        except:
            return None

    @staticmethod
    def extract_urls(file_path):
        """超高效URL提取算法"""
        file_ext = os.path.splitext(file_path)[1].lower()

        try:
            # 文本类文件使用统一处理方法
            if file_ext in ('.txt', '.md', '.html', '.htm', '.xml', '.log'):
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    return UctEngine._extract_text_urls(f.read(), file_path)

            # PDF文件特殊处理
            elif file_ext == '.pdf':
                return UctEngine._extract_pdf_urls(file_path)

            # 其他文件类型使用通用二进制提取
            else:
                with open(file_path, 'rb') as f:
                    content = f.read()
                return UctEngine._extract_binary_urls(content, file_path)

        except Exception as e:
            print(f"提取错误: {file_path} | {str(e)}")
            return []

    @staticmethod
    def _extract_pdf_urls(file_path):
        """PDF URL提取优化"""
        try:
            import fitz  # PyMuPDF
            doc = fitz.open(file_path)
            urls = set()

            for page in doc:
                # 获取所有链接
                links = page.get_links()
                for link in links:
                    if 'uri' in link:
                        urls.add(link['uri'])

                # 获取页面文本
                text = page.get_text("text")
                urls |= set(re.findall(UctConfig.URL_REGEX, text))

            doc.close()
            return UctEngine._normalize_urls(urls, file_path)
        except Exception as e:
            print(f"PDF提取错误: {file_path} | {str(e)}")
            return []

    @staticmethod
    def _extract_text_urls(content, file_path):
        """文本文件URL提取优化"""
        # 两种匹配策略
        urls = set(re.findall(UctConfig.URL_REGEX, content))

        # 第二遍匹配处理边界情况
        alt_urls = set(re.findall(r'https?://[^<\s]{8,}[^\s>]', content))

        return UctEngine._normalize_urls(urls | alt_urls, file_path)

    @staticmethod
    def _extract_binary_urls(content, file_path):
        """通用二进制文件URL提取"""
        try:
            # 尝试解码为文本
            try:
                text = content.decode('utf-8')
            except:
                text = content.decode('latin-1', errors='ignore')

            return UctEngine._extract_text_urls(text, file_path)
        except:
            # 使用正则扫描二进制内容
            text = ''.join(chr(b) if b < 128 and chr(b) in string.printable else '.' for b in content)
            urls = set(re.findall(UctConfig.URL_REGEX, text))
            return UctEngine._normalize_urls(urls, file_path)

    @staticmethod
    def _normalize_urls(urls, file_path):
        """URL标准化处理"""
        results = []
        for url in urls:
            norm_url = UctEngine.normalize_url(url)
            if norm_url:
                results.append({
                    "original_url": url,
                    "normalized_url": norm_url,
                    "source_file": file_path
                })
        return results

    @staticmethod
    def extract_from_text(text, source_name="手动输入"):
        """从文本中提取URL"""
        urls = set()

        # 使用高效的正则提取
        matches = re.findall(UctConfig.URL_REGEX, text)
        for match in matches:
            norm_url = UctEngine.normalize_url(match)
            if norm_url:
                urls.add(norm_url)

        # 准备结果格式
        return [{
            "original_url": url,
            "normalized_url": url,
            "source_file": source_name
        } for url in urls]


# UcT异步验证引擎
class UctVerifier:
    @staticmethod
    async def verify_urls(urls):
        """异步批量验证URL"""
        results = []
        semaphore = asyncio.Semaphore(min(UctConfig.MAX_ASYNC_TASKS, 100))

        async def _verify(url_info):
            async with semaphore:
                return await UctVerifier._verify_url(url_info)

        # 分批验证避免内存溢出
        batch_size = UctConfig.CHUNK_SIZE
        batches = math.ceil(len(urls) / batch_size)

        for i in range(batches):
            start_idx = i * batch_size
            end_idx = min((i + 1) * batch_size, len(urls))
            batch = urls[start_idx:end_idx]

            tasks = [_verify(url_info) for url_info in batch]
            batch_results = await asyncio.gather(*tasks)
            results.extend(batch_results)

        return results

    @staticmethod
    async def _verify_url(url_info):
        """验证单个URL"""
        url = url_info['normalized_url']

        result = {
            **url_info,
            "status_code": 0,
            "status": "未知",
            "emoji": "❓",
            "response_time": 0,
            "error_message": ""
        }

        if not url:
            result['status'] = "无效"
            result['emoji'] = "❌"
            result['error_message'] = "空URL"
            return result

        try:
            # 准备请求参数
            headers = {'User-Agent': UctConfig.USER_AGENT}
            timeout = aiohttp.ClientTimeout(total=UctConfig.REQUEST_TIMEOUT)
            start_time = time.time()

            # 使用异步请求
            async with aiohttp.ClientSession(headers=headers, timeout=timeout) as session:
                # 尝试HEAD请求
                try:
                    async with session.head(url, allow_redirects=True, ssl=False) as response:
                        status = response.status
                        elapsed = time.time() - start_time
                except:
                    # 如果HEAD请求失败，尝试GET请求
                    async with session.get(url, allow_redirects=True, ssl=False) as response:
                        status = response.status
                        elapsed = time.time() - start_time

            # 处理结果
            if status < 200:
                result['status'] = "信息响应"
                result['emoji'] = "ℹ️"
            elif status < 300:
                result['status'] = "活跃"
                result['emoji'] = "✅"
            elif status < 400:
                result['status'] = "重定向"
                result['emoji'] = "🔄"
            elif status < 500:
                result['status'] = "客户端错误"
                result['emoji'] = "⚠️"
            else:
                result['status'] = "服务器错误"
                result['emoji'] = "❌"

            result['status_code'] = status
            result['response_time'] = round(elapsed, 4)
            return result

        except Exception as e:
            return UctVerifier._handle_error(result, e)

    @staticmethod
    def _handle_error(result, exception):
        """错误处理统一入口"""
        err = str(exception).lower()

        if 'timed out' in err:
            result['status'] = "超时"
            result['emoji'] = "⌛"
        elif 'cannot connect' in err:
            result['status'] = "无法连接"
            result['emoji'] = "🔌"
        elif 'name not known' in err or 'gaierror' in err:
            result['status'] = "域名错误"
            result['emoji'] = "🌐"
        elif 'ssl' in err or 'certificate' in err:
            result['status'] = "SSL错误"
            result['emoji'] = "🔒"
        elif 'too many' in err:
            result['status'] = "请求过多"
            result['emoji'] = "🆘"
        else:
            result['status'] = "网络错误"
            result['emoji'] = "❌"

        result['error_message'] = f"{type(exception).__name__}: {str(exception)[:120]}"
        return result


# UcT主界面
class UctApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("UcT - 超高速URL链接检测工具")
        self.geometry("1100x700")
        self.minsize(900, 600)

        # 性能统计
        self.stats = {
            "files": 0,
            "urls_found": 0,
            "urls_verified": 0,
            "success_count": 0,
            "start_time": 0
        }

        # 结果数据
        self.urls = []
        self.results = []
        self.file_paths = []
        self.running = False

        # 创建UI
        self.create_ui()
        self.update_idletasks()

        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        self.status_bar = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # 设置图标
        try:
            self.iconbitmap(self.get_icon_path())
        except:
            pass

    def get_icon_path(self):
        """创建临时图标文件"""
        icon_data = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAB3klEQVQ4T6WT"
            "T0gUYRjGf/PN7Ozfsq2tVdCllIhyCSQpLQpED0EHUYkIgg46lRc9REUE1UHo"
            "WATRoYjCqIsgSFiHEAnpIGEdxDC0NFrX2XV3dj77M7Mf3zs76+64u+khfvAe"
            "5pn3eX+/533fB/wxQaAH+AAkADJ9p2Y+QZ7xr6sQjQIDQC+wQlM2GpJxZQZY"
            "ASbRqB0F9iZ9A8A5wA9sK3ALeN4p5wR9QBR4lM5nQoA7YAYy0QdIqoAlvgLt5Uw+oQFcLpLQFcgGYfPPp+0m/3dYQfzYgvX9l8b9I3Mp7RlMlF2h1aQkKqJ3TgbnRZ69PDB3qA/7XhXaAWx+fvf5o9vQdY2i4i8JqjQvFm9Q0X2GxdK/39LEJwGqD1DxTXX0wfH7"
            "i1vHRfX3hAKqms7GxTV3zE1VK3Hh89+T1U4eJN6x/kIY7Xq2/2X5v/t7x0aG+7cE2hUwUZegGQoEIw/sTlMqVHhWv6RZQq"
            "lQZPrBn8Hh/dG9fmLymR0XXUVQdSXH4cUJjP6dQqYcQBF9yLrBmGzQ3t4cQrI4PcGp67k0q9S2W8RkGUiLw7cUfTpz0+QL09m6j2"
            "XSxLRfXkTR1g5kf6w+HTj9Whge4o5ZKN43hsb1Y5lL7Jpt5mNVdGq6YQpouVq6MlD7B9eYpKuYZWYr4ZfPv+1GZbR6C0VcT3O0t6sYyV17EzdSRZRDQ8pKqDuVHEq9p4ioXQ3UM4Gm7X4A6QByZq9Tqp5RwBz0MSBFwJbEXFytbx8yauYuLJHlIkQvjSORLjZxBiUXx+P/90IQP8BrJAI3rD4T0eAAAAAElFTkSuQmCC"
        )

        _, icon_path = tempfile.mkstemp(suffix='.ico')
        with open(icon_path, 'wb') as icon_file:
            icon_file.write(icon_data)

        return icon_path

    def create_ui(self):
        """创建主界面"""
        # 主容器
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 标题栏
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(title_frame, text="UcT", font=("Arial", 24, "bold"), foreground="#3498db").pack(side=tk.LEFT)
        ttk.Label(title_frame, text="超高速URL链接检测工具", font=("Arial", 14)).pack(side=tk.LEFT, padx=10)

        # =================== URL输入面板 ===================
        input_frame = ttk.LabelFrame(main_frame, text="手动输入URL")
        input_frame.pack(fill=tk.X, pady=(0, 10))

        # URL输入框
        ttk.Label(input_frame, text="直接输入URL（每行一个，或多个用空格分隔）:").pack(anchor=tk.W, padx=5, pady=5)

        self.url_input = scrolledtext.ScrolledText(input_frame, height=5)
        self.url_input.pack(fill=tk.X, padx=5, pady=(0, 5))

        # 输入控制按钮
        input_btn_frame = ttk.Frame(input_frame)
        input_btn_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(
            input_btn_frame,
            text="从剪贴板导入",
            command=self.paste_from_clipboard
        ).pack(side=tk.LEFT, padx=(0, 5))

        ttk.Button(
            input_btn_frame,
            text="清空输入框",
            command=self.clear_input
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            input_btn_frame,
            text="添加URL并立即检查",
            command=self.check_input_urls
        ).pack(side=tk.RIGHT)

        # =================== 控制面板 ===================
        control_frame = ttk.LabelFrame(main_frame, text="控制面板")
        control_frame.pack(fill=tk.X, pady=(0, 10))

        # 按钮组
        btn_frame = ttk.Frame(control_frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)

        self.add_file_btn = ttk.Button(btn_frame, text="添加文件", width=12, command=self.add_files)
        self.add_file_btn.pack(side=tk.LEFT, padx=(0, 5))

        self.add_folder_btn = ttk.Button(btn_frame, text="添加文件夹", width=12, command=self.add_folder)
        self.add_folder_btn.pack(side=tk.LEFT, padx=5)

        self.start_btn = ttk.Button(btn_frame, text="开始检测", width=12, command=self.start_validation)
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.stop_btn = ttk.Button(btn_frame, text="停止", width=12, command=self.stop_validation, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        self.export_btn = ttk.Button(btn_frame, text="导出结果", width=12, command=self.export_results, state=tk.DISABLED)
        self.export_btn.pack(side=tk.RIGHT)

        # =================== 文件列表 ===================
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.X, pady=(0, 10))

        scroll_y = ttk.Scrollbar(list_frame, orient=tk.VERTICAL)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        scroll_x = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.file_list = tk.Listbox(
            list_frame,
            height=5,
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set
        )
        self.file_list.pack(fill=tk.BOTH, expand=True)

        scroll_y.config(command=self.file_list.yview)
        scroll_x.config(command=self.file_list.xview)

        # 右键菜单
        self.list_menu = tk.Menu(self, tearoff=0)
        self.list_menu.add_command(label="删除选中项", command=self.remove_selected)
        self.file_list.bind("<Button-3>", self.show_list_menu)

        # =================== 状态面板 ===================
        status_frame = ttk.LabelFrame(main_frame, text="进度状态")
        status_frame.pack(fill=tk.X, pady=(0, 10))

        # 进度条
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_bar = ttk.Progressbar(status_frame, length=100, mode='determinate', variable=self.progress_var)
        self.progress_bar.pack(fill=tk.X, padx=5, pady=(0, 5))

        # 统计信息
        stats_frame = ttk.Frame(status_frame)
        stats_frame.pack(fill=tk.X, padx=5, pady=5)

        labels = ["文件", "URL发现", "已验证", "有效链接", "时间"]
        self.stats_vars = {label: tk.StringVar(value="0") for label in labels}

        for i, label in enumerate(labels):
            frame = ttk.Frame(stats_frame)
            frame.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            ttk.Label(frame, text=label + ":", font=("Arial", 9)).pack(anchor=tk.W)
            ttk.Label(frame, textvariable=self.stats_vars[label], font=("Arial", 10, "bold")).pack(anchor=tk.W)

        # =================== 结果展示 ===================
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)

        # 结果表格
        results_frame = ttk.Frame(notebook)
        notebook.add(results_frame, text="URL结果")

        # 结果列定义
        columns = [
            ("状态", 50),
            ("URL", 450),
            ("状态码", 80),
            ("响应时间", 80),
            ("来源", 150)
        ]

        self.results_tree = ttk.Treeview(
            results_frame,
            columns=[col[0] for col in columns],
            show='headings',
            selectmode='browse'
        )

        # 配置列
        for col_name, width in columns:
            self.results_tree.heading(col_name, text=col_name)
            self.results_tree.column(col_name, width=width, stretch=False)

        # 滚动条
        tree_scroll_y = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.results_tree.yview)
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.results_tree.configure(yscrollcommand=tree_scroll_y.set)

        tree_scroll_x = ttk.Scrollbar(results_frame, orient=tk.HORIZONTAL, command=self.results_tree.xview)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.results_tree.configure(xscrollcommand=tree_scroll_x.set)

        self.results_tree.pack(fill=tk.BOTH, expand=True)
        self.results_tree.bind("<Double-1>", self.on_result_double_click)

        # 表格右键菜单
        self.tree_menu = tk.Menu(self, tearoff=0)
        self.tree_menu.add_command(label="打开链接", command=self.open_selected_url)
        self.tree_menu.add_command(label="复制链接", command=self.copy_selected_url)
        self.tree_menu.add_separator()
        self.tree_menu.add_command(label="删除选中项", command=self.delete_selected_result)
        self.results_tree.bind("<Button-3>", self.show_tree_menu)

        # 报告面板
        report_frame = ttk.Frame(notebook)
        notebook.add(report_frame, text="分析报告")

        self.report_text = scrolledtext.ScrolledText(report_frame, wrap=tk.WORD)
        self.report_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.report_text.config(state=tk.DISABLED)

    def paste_from_clipboard(self):
        """从剪贴板粘贴内容到输入框"""
        try:
            clipboard_text = self.clipboard_get()
            if clipboard_text.strip():
                self.url_input.delete(1.0, tk.END)
                self.url_input.insert(tk.END, clipboard_text)
                self.status_var.set("已从剪贴板导入URL")
        except:
            self.status_var.set("无法读取剪贴板内容")

    def clear_input(self):
        """清空URL输入框"""
        self.url_input.delete(1.0, tk.END)
        self.status_var.set("已清空URL输入框")

    def check_input_urls(self):
        """直接检查输入的URL"""
        input_text = self.url_input.get(1.0, tk.END).strip()
        if not input_text:
            self.status_var.set("URL输入框为空")
            return

        # 提取URL
        urls = UctEngine.extract_from_text(input_text)
        if not urls:
            self.status_var.set("未发现有效URL")
            return

        # 重置状态为仅检查输入
        self.start_validation(only_manual=True)

    def add_files(self):
        """添加文件"""
        file_types = [
            ('所有文件', '*.*'),
            ('文档文件', '*.docx;*.doc;*.pdf;*.pptx;*.ppt'),
            ('电子表格', '*.xlsx;*.xls'),
            ('网页文件', '*.html;*.htm'),
            ('文本文件', '*.txt;*.md')
        ]

        files = filedialog.askopenfilenames(filetypes=file_types)
        if files:
            self.update_file_list(files)

    def add_folder(self):
        """添加文件夹"""
        folder = filedialog.askdirectory(title="选择文件夹")
        if folder:
            # 多线程收集文件
            def collect_files():
                files = []
                for root, _, filenames in os.walk(folder):
                    for fn in filenames:
                        if '.' in fn and not fn.startswith('~'):  # 忽略临时文件
                            files.append(os.path.join(root, fn))
                return files

            threading.Thread(target=lambda: self.update_file_list(collect_files())).start()

    def update_file_list(self, files):
        """更新文件列表"""
        new_files = [f for f in files if f not in self.file_paths]

        if new_files:
            self.file_paths.extend(new_files)

            for f in new_files:
                self.file_list.insert(tk.END, f"● {os.path.basename(f)}")

            self.stats_vars["文件"].set(str(len(self.file_paths)))
            self.status_var.set(f"添加了 {len(new_files)} 个文件")

    def show_list_menu(self, event):
        """显示文件列表右键菜单"""
        try:
            if self.file_list.curselection():
                self.list_menu.post(event.x_root, event.y_root)
        except:
            pass

    def remove_selected(self):
        """删除选中的文件项"""
        if not self.running:
            selected = self.file_list.curselection()
            for index in selected[::-1]:
                path = self.file_paths[index]
                del self.file_paths[index]
                self.file_list.delete(index)

            self.stats_vars["文件"].set(str(len(self.file_paths)))
            self.status_var.set(f"已删除 {len(selected)} 个文件")

    def start_validation(self, only_manual=False):
        """开始检测"""
        if only_manual:
            # 仅处理手动输入的URL
            if self.running:
                return

            # 重置状态
            self.running = True
            self.urls.clear()
            self.results.clear()

            # UI更新
            self.start_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk.NORMAL)
            self.export_btn.config(state=tk.DISABLED)

            # 从输入框获取URL
            input_text = self.url_input.get(1.0, tk.END).strip()
            if input_text:
                self.urls = UctEngine.extract_from_text(input_text)

            if not self.urls:
                messagebox.showinfo("提示", "未找到有效URL")
                self.finish_validation()
                return

            # 初始化统计
            self.stats = {
                "files": 0,
                "urls_found": len(self.urls),
                "urls_verified": 0,
                "success_count": 0,
                "start_time": time.time()
            }

            # 更新统计显示
            self.stats_vars["文件"].set("0")
            self.stats_vars["URL发现"].set(str(len(self.urls)))
            self.stats_vars["已验证"].set("0")
            self.stats_vars["有效链接"].set("0")
            self.stats_vars["时间"].set("0")

            self.progress_var.set(0)
            self.status_var.set("开始验证手动输入的URL...")

            # 清空结果树
            self.results_tree.delete(*self.results_tree.get_children())
            self.report_text.config(state=tk.NORMAL)
            self.report_text.delete(1.0, tk.END)
            self.report_text.config(state=tk.DISABLED)

            # 在后台线程中处理
            threading.Thread(target=self.verify_urls_only, daemon=True).start()
        else:
            # 处理文件+手动输入
            if not self.file_paths:
                input_text = self.url_input.get(1.0, tk.END).strip()
                if not input_text:
                    messagebox.showinfo("提示", "请添加文档或文件夹，或输入URL")
                    return

            if self.running:
                return

            # 重置状态
            self.running = True
            self.urls.clear()
            self.results.clear()

            # UI更新
            self.start_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk.NORMAL)
            self.export_btn.config(state=tk.DISABLED)
            self.file_list.delete(0, tk.END)

            for path in self.file_paths:
                self.file_list.insert(tk.END, f"● {os.path.basename(path)}")

            # 清空结果
            self.results_tree.delete(*self.results_tree.get_children())
            self.report_text.config(state=tk.NORMAL)
            self.report_text.delete(1.0, tk.END)
            self.report_text.config(state=tk.DISABLED)

            # 初始化统计
            self.stats = {
                "files": len(self.file_paths),
                "urls_found": 0,
                "urls_verified": 0,
                "success_count": 0,
                "start_time": time.time()
            }

            # 更新统计显示
            for key in self.stats_vars:
                self.stats_vars[key].set("0")

            self.progress_var.set(0)
            self.status_var.set("开始处理...")

            # 在后台线程中处理
            threading.Thread(target=self.process_files, daemon=True).start()

    def verify_urls_only(self):
        """仅验证URL（不处理文件）"""
        # 在事件循环中运行异步验证
        asyncio.run(self.verify_urls_async())

        # 完成验证
        self.finish_validation()

    def stop_validation(self):
        """停止检测"""
        if self.running:
            self.running = False
            self.status_var.set("操作已停止")
            self.stop_btn.config(state=tk.DISABLED)

    def process_files(self):
        """处理文件"""
        # 阶段1：提取URL
        self.status_var.set("正在提取URL...")

        # 使用线程池并行处理文件
        with concurrent.futures.ThreadPoolExecutor(max_workers=UctConfig.MAX_WORKERS) as executor:
            futures = [executor.submit(UctEngine.extract_urls, fp) for fp in self.file_paths]

            for i, future in enumerate(concurrent.futures.as_completed(futures)):
                if not self.running:
                    break

                urls = future.result()
                if urls:
                    self.urls.extend(urls)
                    self.stats["urls_found"] += len(urls)
                    self.stats_vars["URL发现"].set(str(self.stats["urls_found"]))

                # 更新进度
                progress = (i + 1) / len(self.file_paths) * 50
                self.progress_var.set(progress)
                self.status_var.set(f"文件处理中 {i + 1}/{len(self.file_paths)}")

        # 添加手动输入的URL
        if self.running:
            input_text = self.url_input.get(1.0, tk.END).strip()
            if input_text:
                manual_urls = UctEngine.extract_from_text(input_text)
                if manual_urls:
                    self.urls.extend(manual_urls)
                    self.stats["urls_found"] += len(manual_urls)
                    self.stats_vars["URL发现"].set(str(self.stats["urls_found"]))
                    self.status_var.set(f"添加了 {len(manual_urls)} 个手动输入的URL")

        # 如果被停止或没有找到URL
        if not self.running or not self.urls:
            self.finish_validation()
            return

        # 阶段2：验证URL
        self.status_var.set("正在验证URL...")

        # 在事件循环中运行异步验证
        asyncio.run(self.verify_urls_async())

        # 完成验证
        self.finish_validation()

    async def verify_urls_async(self):
        """异步验证URL"""
        # 批量处理URL验证
        results = await UctVerifier.verify_urls(self.urls)

        # 分批更新UI，避免界面冻结
        batch_size = UctConfig.CHUNK_SIZE
        total = len(results)

        for i in range(0, total, batch_size):
            if not self.running:
                break

            batch = results[i:i + batch_size]
            self.process_results_batch(batch)

            # 更新进度
            if self.stats["files"] > 0:
                # 文件+输入混合模式
                progress = 50 + (i / total * 50)
            else:
                # 纯手动输入模式
                progress = (i / total * 100)

            self.progress_var.set(min(progress, 100))
            self.stats_vars["已验证"].set(str(self.stats["urls_verified"]))

            # 避免过于频繁的更新
            time.sleep(0.05)

        # 保存最终结果
        self.results = results

    def process_results_batch(self, batch):
        """处理一批结果"""
        # 更新统计
        self.stats["urls_verified"] += len(batch)
        self.stats["success_count"] += sum(
            1 for res in batch
            if res.get('status') in ("活跃", "重定向")
        )

        # 更新成功计数
        self.stats_vars["有效链接"].set(str(self.stats["success_count"]))

        # 添加结果到Treeview
        for result in batch:
            self.add_result_to_tree(result)

        # 添加后自动滚动到底部
        self.results_tree.yview_moveto(1.0)

    def add_result_to_tree(self, result):
        """添加结果到Treeview"""
        tags = ""
        if result['status'] == "活跃":
            tags = "success"
        elif result['status'] in ("客户端错误", "服务器错误"):
            tags = "error"
        elif result['status'] == "超时":
            tags = "warning"

        # 添加行到Treeview
        self.results_tree.insert("", "end", values=(
            result.get('emoji', '❓'),
            result['normalized_url'][:100] + ('...' if len(result['normalized_url']) > 100 else ''),
            result.get('status_code', 'N/A'),
            f"{result.get('response_time', 0):.3f}s",
            os.path.basename(result['source_file']) if len(result['source_file']) < 30 else result['source_file'][:27] + "..."
        ), tags=(tags,))

        # 配置标签样式（只需一次）
        if not hasattr(self, 'tags_configured'):
            self.results_tree.tag_configure('success', foreground='#2ecc71')
            self.results_tree.tag_configure('error', foreground='#e74c3c')
            self.results_tree.tag_configure('warning', foreground='#f39c12')
            self.tags_configured = True

    def finish_validation(self):
        """完成验证过程"""
        self.running = False

        # 计算耗时
        elapsed = time.time() - self.stats["start_time"]
        minutes, seconds = divmod(elapsed, 60)
        time_str = f"{int(minutes)}分{int(seconds)}秒"
        self.stats_vars["时间"].set(time_str)

        # 更新UI状态
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.export_btn.config(state=tk.NORMAL)

        # 更新状态
        self.progress_var.set(100)
        self.status_var.set("检测完成!")

        # 生成报告
        self.generate_report()

    def generate_report(self):
        """生成分析报告"""
        if not self.results:
            return

        # 收集统计信息
        report = "UcT 检测报告\n"
        report += "=" * 50 + "\n\n"
        report += f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        report += f"处理文件: {self.stats['files']} 个\n"
        report += f"发现链接: {self.stats['urls_found']} 个\n"
        report += f"验证链接: {self.stats['urls_verified']} 个\n"
        report += f"有效链接: {self.stats['success_count']} 个 (成功率: {self.stats['success_count'] / max(self.stats['urls_verified'], 1) * 100:.1f}%)\n"
        report += f"耗时: {self.stats_vars['时间'].get()}\n\n"

        # 状态分布
        status_stats = defaultdict(int)
        for res in self.results:
            status_stats[res['status']] += 1

        if status_stats:
            report += "状态分布:\n"
            report += "--------\n"
            for status, count in sorted(status_stats.items(), key=lambda x: x[1], reverse=True):
                emoji = {
                    "活跃": "✅",
                    "重定向": "🔄",
                    "客户端错误": "⚠️",
                    "服务器错误": "❌",
                    "超时": "⌛"
                }.get(status, " ")
                report += f"{emoji} {status:<12} {count} 个\n"
            report += "\n"

        # 更新报告文本框
        self.report_text.config(state=tk.NORMAL)
        self.report_text.delete(1.0, tk.END)
        self.report_text.insert(tk.END, report)
        self.report_text.config(state=tk.DISABLED)

    def export_results(self):
        """导出结果"""
        if not self.results:
            return

        file_path = filedialog.asksaveasfilename(
            title="导出检测结果",
            filetypes=[("CSV文件", "*.csv"), ("Excel文件", "*.xlsx")],
            defaultextension=".csv"
        )

        if not file_path:
            return

        try:
            # 准备数据
            data = []
            for res in self.results:
                data.append({
                    "URL": res['normalized_url'],
                    "状态": res['status'],
                    "状态码": res.get('status_code', ''),
                    "响应时间": res.get('response_time', ''),
                    "错误信息": res.get('error_message', ''),
                    "来源文件": res['source_file']
                })

            df = pd.DataFrame(data)

            # 保存为CSV
            if file_path.endswith('.csv'):
                df.to_csv(file_path, index=False, encoding='utf-8-sig')

            # 保存为Excel
            elif file_path.endswith('.xlsx'):
                df.to_excel(file_path, index=False)

            self.status_var.set(f"结果已导出: {os.path.basename(file_path)}")
            messagebox.showinfo("导出成功", f"结果已保存到:\n{file_path}")

        except Exception as e:
            messagebox.showerror("导出错误", f"导出失败: {str(e)}")

    def on_result_double_click(self, event):
        """双击打开URL"""
        self.open_selected_url()

    def show_tree_menu(self, event):
        """显示结果树的右键菜单"""
        if self.results_tree.identify_row(event.y) and self.results_tree.selection():
            self.tree_menu.post(event.x_root, event.y_root)

    def open_selected_url(self):
        """打开选中的URL"""
        selected = self.results_tree.selection()
        if not selected:
            return

        item = self.results_tree.focus()
        values = self.results_tree.item(item, 'values')
        if not values or len(values) < 2:
            return

        url = values[1]  # 第二列是URL
        if not url.startswith('http'):
            return

        # 查找完整URL
        full_url = next((r['normalized_url'] for r in self.results if r['normalized_url'].startswith(url)), "")

        if full_url:
            try:
                webbrowser.open(full_url)
                self.status_var.set(f"打开: {full_url[:60]}...")
            except:
                messagebox.showerror("错误", "无法打开URL")

    def copy_selected_url(self):
        """复制选中的URL"""
        selected = self.results_tree.selection()
        if not selected:
            return

        item = self.results_tree.focus()
        values = self.results_tree.item(item, 'values')
        if not values or len(values) < 2:
            return

        url = values[1]  # 第二列是URL

        # 查找完整URL
        full_url = next((r['normalized_url'] for r in self.results if r['normalized_url'].startswith(url)), "")

        if full_url:
            try:
                self.clipboard_clear()
                self.clipboard_append(full_url)
                self.status_var.set(f"已复制URL: {full_url[:40]}...")
            except:
                messagebox.showerror("错误", "无法复制URL")

    def delete_selected_result(self):
        """删除选中的结果项"""
        selected = self.results_tree.selection()
        if not selected:
            return

        # 从Treeview删除
        for item in selected:
            self.results_tree.delete(item)

        # 更新结果列表
        self.results = [res for res in self.results if res['normalized_url'] not in
                        [self.results_tree.item(item, 'values')[1] for item in selected]]

        self.status_var.set(f"已删除 {len(selected)} 个结果项")

        # 重新生成报告
        self.generate_report()


def main():
    """主入口函数"""
    # 静默检查并安装依赖
    if not SilentDependencyManager.check_and_install_dependencies():
        messagebox.showerror(
            "依赖安装失败",
            "无法自动安装必要的依赖库。\n\n"
            "请手动安装以下依赖：\n\n"
            f"{', '.join(UctConfig.REQUIRED_PACKAGES)}\n\n"
            "使用命令：\n"
            f"pip install {' '.join(UctConfig.REQUIRED_PACKAGES)}"
        )
        sys.exit(1)

    # 启动主程序
    app = UctApp()
    app.mainloop()


if __name__ == "__main__":
    main()