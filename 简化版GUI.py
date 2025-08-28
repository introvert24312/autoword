#!/usr/bin/env python3
"""
AutoWord vNext 简化版GUI
使用最少依赖的图形界面
"""

import sys
import os
import json
import time
import threading
from pathlib import Path
from typing import Optional, Dict, Any

# 添加项目路径
sys.path.insert(0, str(Path(__file__).parent))

try:
    from tkinter import *
    from tkinter import ttk, filedialog, messagebox, scrolledtext
    from tkinter.ttk import Progressbar
except ImportError:
    print("❌ Tkinter未安装，这是Python的标准库，请检查Python安装")
    sys.exit(1)

class AutoWordGUI:
    def __init__(self):
        self.root = Tk()
        self.setup_window()
        self.setup_variables()
        self.setup_ui()
        self.load_config()
        
    def setup_window(self):
        """设置主窗口"""
        self.root.title("AutoWord vNext - 智能文档处理系统")
        self.root.geometry("800x700")
        self.root.resizable(True, True)
        
        # 设置图标（如果存在）
        try:
            self.root.iconbitmap("autoword.ico")
        except:
            pass
            
        # 设置样式
        style = ttk.Style()
        style.theme_use('clam')
        
    def setup_variables(self):
        """设置变量"""
        self.api_provider = StringVar(value="openai")
        self.api_key = StringVar()
        self.input_file = StringVar()
        self.output_dir = StringVar()
        self.user_intent = StringVar()
        self.progress_var = DoubleVar()
        self.status_var = StringVar(value="就绪")
        self.processing = False
        
        # 批注处理变量
        self.extracted_comments = []
        self.comments_json = None
        
        # API配置存储
        self.backup_config = None
        self.openai_key = ""
        self.claude_key = ""
        
        # 封面保护变量
        self.cover_protection_enabled = True
        
    def setup_ui(self):
        """设置用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(W, E, N, S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        row = 0
        
        # 标题
        title_label = ttk.Label(main_frame, text="AutoWord vNext", font=("Arial", 16, "bold"))
        title_label.grid(row=row, column=0, columnspan=3, pady=(0, 20))
        row += 1
        
        # API配置区域
        config_frame = ttk.LabelFrame(main_frame, text="AI模型配置", padding="10")
        config_frame.grid(row=row, column=0, columnspan=3, sticky=(W, E), pady=(0, 10))
        config_frame.columnconfigure(1, weight=1)
        row += 1
        
        # 模型选择
        ttk.Label(config_frame, text="AI模型:").grid(row=0, column=0, sticky=W, padx=(0, 10))
        model_combo = ttk.Combobox(config_frame, textvariable=self.api_provider, 
                                  values=["openai", "anthropic"], state="readonly", width=15)
        model_combo.grid(row=0, column=1, sticky=W, padx=(0, 10))
        model_combo.bind('<<ComboboxSelected>>', self.on_model_change)
        
        # API密钥
        ttk.Label(config_frame, text="API密钥:").grid(row=0, column=2, sticky=W, padx=(20, 10))
        api_key_entry = ttk.Entry(config_frame, textvariable=self.api_key, show="*", width=30)
        api_key_entry.grid(row=0, column=3, sticky=(W, E), padx=(0, 10))
        
        # 显示/隐藏密钥按钮
        self.show_key_var = BooleanVar()
        show_key_btn = ttk.Checkbutton(config_frame, text="显示", variable=self.show_key_var,
                                      command=self.toggle_key_visibility)
        show_key_btn.grid(row=0, column=4, sticky=W)
        
        self.api_key_entry = api_key_entry
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="10")
        file_frame.grid(row=row, column=0, columnspan=3, sticky=(W, E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        row += 1
        
        # 输入文件
        ttk.Label(file_frame, text="输入文件:").grid(row=0, column=0, sticky=W, padx=(0, 10))
        ttk.Entry(file_frame, textvariable=self.input_file, state="readonly").grid(row=0, column=1, sticky=(W, E), padx=(0, 10))
        ttk.Button(file_frame, text="选择文件", command=self.select_input_file).grid(row=0, column=2, sticky=W)
        
        # 输出目录
        ttk.Label(file_frame, text="输出目录:").grid(row=1, column=0, sticky=W, padx=(0, 10), pady=(10, 0))
        ttk.Entry(file_frame, textvariable=self.output_dir, state="readonly").grid(row=1, column=1, sticky=(W, E), padx=(0, 10), pady=(10, 0))
        ttk.Button(file_frame, text="选择目录", command=self.select_output_dir).grid(row=1, column=2, sticky=W, pady=(10, 0))
        
        # 批注处理区域
        comment_frame = ttk.LabelFrame(main_frame, text="批注处理", padding="10")
        comment_frame.grid(row=row, column=0, columnspan=3, sticky=(W, E), pady=(0, 10))
        comment_frame.columnconfigure(1, weight=1)
        row += 1
        
        # 批注处理选项
        self.use_comments = BooleanVar(value=True)
        ttk.Checkbutton(comment_frame, text="读取批注作为指令", variable=self.use_comments,
                       command=self.toggle_comment_processing).grid(row=0, column=0, sticky=W, padx=(0, 20))
        
        self.execute_tags_only = BooleanVar(value=False)
        ttk.Checkbutton(comment_frame, text="只执行带EXECUTE标签的批注", 
                       variable=self.execute_tags_only).grid(row=0, column=1, sticky=W, padx=(0, 20))
        
        self.llm_fallback = BooleanVar(value=True)
        ttk.Checkbutton(comment_frame, text="批注解析失败时使用LLM", 
                       variable=self.llm_fallback).grid(row=0, column=2, sticky=W)
        
        # 封面保护选项（第二行）
        self.cover_protection = BooleanVar(value=True)
        ttk.Checkbutton(comment_frame, text="🛡️ 保护封面和目录格式", variable=self.cover_protection,
                       command=self.toggle_cover_protection).grid(row=1, column=0, sticky=W, padx=(0, 20), pady=(5, 0))
        
        self.auto_page_break = BooleanVar(value=False)
        ttk.Checkbutton(comment_frame, text="📄 自动插入分页符", variable=self.auto_page_break).grid(
            row=1, column=1, sticky=W, padx=(0, 20), pady=(5, 0))
        
        # 封面保护状态显示
        self.cover_status = StringVar(value="封面保护: 启用")
        ttk.Label(comment_frame, textvariable=self.cover_status, foreground="green", font=("Arial", 9)).grid(
            row=1, column=2, sticky=W, pady=(5, 0))
        
        # 批注状态显示
        self.comments_status = StringVar(value="未检测到批注")
        ttk.Label(comment_frame, textvariable=self.comments_status, foreground="gray").grid(
            row=2, column=0, columnspan=3, sticky=W, pady=(5, 0))
        
        # 用户意图区域
        intent_frame = ttk.LabelFrame(main_frame, text="处理指令", padding="10")
        intent_frame.grid(row=row, column=0, columnspan=3, sticky=(W, E), pady=(0, 10))
        intent_frame.columnconfigure(0, weight=1)
        row += 1
        
        # 处理指令输入
        ttk.Label(intent_frame, text="附加处理指令 (可选):").grid(row=0, column=0, sticky=W, pady=(0, 5))
        
        # 说明文字
        info_label = ttk.Label(intent_frame, text="💡 系统会自动从文档批注中提取处理指令，此处可添加额外指令", 
                              foreground="gray", font=("Arial", 9))
        info_label.grid(row=1, column=0, sticky=W, pady=(0, 5))
        
        intent_text = Text(intent_frame, height=3, wrap=WORD)
        intent_text.grid(row=2, column=0, sticky=(W, E), pady=(0, 10))
        intent_text.bind('<KeyRelease>', self.update_intent_from_text)
        self.intent_text = intent_text
        
        # 控制按钮区域
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=row, column=0, columnspan=3, pady=(0, 10))
        row += 1
        
        self.process_btn = ttk.Button(control_frame, text="🚀 开始处理", command=self.start_processing)
        self.process_btn.grid(row=0, column=0, padx=(0, 10))
        
        self.stop_btn = ttk.Button(control_frame, text="⏹️ 停止", command=self.stop_processing, state=DISABLED)
        self.stop_btn.grid(row=0, column=1, padx=(0, 10))
        
        self.dry_run_btn = ttk.Button(control_frame, text="🔍 预览批注", command=self.preview_comments, state=DISABLED)
        self.dry_run_btn.grid(row=0, column=2, padx=(0, 10))
        
        ttk.Button(control_frame, text="💾 保存配置", command=self.save_config).grid(row=0, column=3, padx=(0, 10))
        ttk.Button(control_frame, text="🛡️ 测试封面保护", command=self.test_cover_protection).grid(row=0, column=4, padx=(0, 10))
        ttk.Button(control_frame, text="🧪 性能测试", command=self.run_performance_test).grid(row=0, column=5)
        
        # 进度区域
        progress_frame = ttk.LabelFrame(main_frame, text="处理进度", padding="10")
        progress_frame.grid(row=row, column=0, columnspan=3, sticky=(W, E), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        row += 1
        
        # 进度条
        self.progress_bar = Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=0, column=0, sticky=(W, E), pady=(0, 5))
        
        # 状态标签
        self.status_label = ttk.Label(progress_frame, textvariable=self.status_var)
        self.status_label.grid(row=1, column=0, sticky=W)
        
        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="处理日志", padding="10")
        log_frame.grid(row=row, column=0, columnspan=3, sticky=(W, E, N, S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(row, weight=1)
        
        # 日志文本区域
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, wrap=WORD)
        self.log_text.grid(row=0, column=0, sticky=(W, E, N, S))
        
        # 日志控制按钮
        log_btn_frame = ttk.Frame(log_frame)
        log_btn_frame.grid(row=1, column=0, sticky=E, pady=(5, 0))
        
        ttk.Button(log_btn_frame, text="清空日志", command=self.clear_log).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(log_btn_frame, text="保存日志", command=self.save_log).grid(row=0, column=1)
        
    def toggle_key_visibility(self):
        """切换API密钥显示/隐藏"""
        if self.show_key_var.get():
            self.api_key_entry.config(show="")
        else:
            self.api_key_entry.config(show="*")
    
    def on_model_change(self, event=None):
        """处理模型切换"""
        try:
            current_provider = self.api_provider.get()
            
            if current_provider == "openai":
                # 切换到OpenAI
                if self.openai_key:
                    self.api_key.set(self.openai_key)
                    self.log_message("🤖 已切换到 OpenAI GPT-4")
                else:
                    self.log_message("⚠️ OpenAI密钥未配置")
                    
            elif current_provider == "anthropic":
                # 切换到Claude
                if self.claude_key:
                    self.api_key.set(self.claude_key)
                    self.log_message("🧠 已切换到 Anthropic Claude")
                else:
                    self.log_message("⚠️ Claude密钥未配置")
            
        except Exception as e:
            self.log_message(f"⚠️ 模型切换失败: {e}")
    
    def select_input_file(self):
        """选择输入文件"""
        filename = filedialog.askopenfilename(
            title="选择Word文档",
            filetypes=[("Word文档", "*.docx *.doc"), ("所有文件", "*.*")]
        )
        if filename:
            self.input_file.set(filename)
            # 自动设置输出目录
            if not self.output_dir.get():
                self.output_dir.set(str(Path(filename).parent))
            
            # 自动从文档批注中提取处理指令
            self.extract_comments_from_document(filename)
    
    def select_output_dir(self):
        """选择输出目录"""
        dirname = filedialog.askdirectory(title="选择输出目录")
        if dirname:
            self.output_dir.set(dirname)
    
    def update_intent_from_text(self, event=None):
        """从文本框更新用户意图"""
        content = self.intent_text.get("1.0", END).strip()
        self.user_intent.set(content)
    
    def log_message(self, message: str, level: str = "INFO"):
        """添加日志消息"""
        timestamp = time.strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {level}: {message}\n"
        
        self.log_text.insert(END, log_entry)
        self.log_text.see(END)
        self.root.update_idletasks()
    
    def clear_log(self):
        """清空日志"""
        self.log_text.delete("1.0", END)
    
    def save_log(self):
        """保存日志"""
        filename = filedialog.asksaveasfilename(
            title="保存日志",
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.get("1.0", END))
                self.log_message(f"日志已保存到: {filename}")
            except Exception as e:
                messagebox.showerror("错误", f"保存日志失败: {e}")
    
    def validate_inputs(self) -> bool:
        """验证输入"""
        if not self.api_key.get().strip():
            messagebox.showerror("错误", "请输入API密钥")
            return False
        
        if not self.input_file.get():
            messagebox.showerror("错误", "请选择输入文件")
            return False
        
        if not Path(self.input_file.get()).exists():
            messagebox.showerror("错误", "输入文件不存在")
            return False
        
        intent = self.user_intent.get().strip()
        if not intent:
            intent = self.intent_text.get("1.0", END).strip()
        
        if not intent:
            messagebox.showerror("错误", "请输入处理指令")
            return False
        
        return True
    
    def start_processing(self):
        """开始处理"""
        if not self.validate_inputs():
            return
        
        if self.processing:
            messagebox.showwarning("警告", "正在处理中，请等待完成")
            return
        
        # 更新UI状态
        self.processing = True
        self.process_btn.config(state=DISABLED)
        self.stop_btn.config(state=NORMAL)
        self.progress_var.set(0)
        self.status_var.set("正在初始化...")
        
        # 获取处理参数
        intent = self.user_intent.get().strip()
        if not intent:
            intent = self.intent_text.get("1.0", END).strip()
        
        # 在新线程中处理
        thread = threading.Thread(target=self.process_document, args=(intent,))
        thread.daemon = True
        thread.start()
    
    def process_document(self, user_intent: str):
        """处理文档（在后台线程中运行）"""
        try:
            self.log_message("开始处理文档...")
            self.progress_var.set(10)
            self.status_var.set("正在加载AutoWord...")
            
            # 导入AutoWord（使用新的SimplePipeline）
            try:
                from autoword.vnext.simple_pipeline import SimplePipeline
                from autoword.vnext.core import VNextConfig, LLMConfig
                use_simple_pipeline = True
                self.log_message("🔧 使用增强版SimplePipeline（支持封面保护）")
            except ImportError:
                # 回退到原版本
                from autoword.vnext import VNextPipeline, VNextConfig, LLMConfig
                use_simple_pipeline = False
                self.log_message("⚠️ 使用标准版Pipeline")
            
            self.progress_var.set(20)
            self.status_var.set("正在配置AI模型...")
            
            # 创建配置
            provider = self.api_provider.get()
            api_key = self.api_key.get().strip()
            
            # 根据提供商设置模型和基础URL
            if provider == "openai":
                model = "gpt-4"
                base_url = "https://globalai.vip"
            else:  # anthropic
                model = "claude-3-sonnet-20240229"  # 修正模型名称
                base_url = "https://globalai.vip"
            
            # 根据使用的Pipeline类型创建配置
            if use_simple_pipeline:
                # 使用SimplePipeline（支持封面保护）
                llm_config = LLMConfig(
                    provider=provider,
                    model=model,
                    api_key=api_key,
                    temperature=0.1
                )
                config = VNextConfig(llm=llm_config)
                pipeline = SimplePipeline(config)
                self.log_message("✅ 已启用封面保护功能")
            else:
                # 使用原版Pipeline
                if provider == "openai":
                    llm_config = LLMConfig(
                        provider="openai",
                        model="gpt-4",
                        api_key=api_key,
                        temperature=0.1
                    )
                else:  # anthropic
                    llm_config = LLMConfig(
                        provider="anthropic", 
                        model="claude-3-sonnet-20240229",
                        api_key=api_key,
                        temperature=0.1
                    )
                
                config = VNextConfig(llm=llm_config)
                pipeline = VNextPipeline(config)
            
            # 添加批注处理配置（作为额外属性）
            if hasattr(config, 'comment_processing') or not use_simple_pipeline:
                config.comment_processing = {
                    "enabled": self.use_comments.get(),
                    "execute_tags_only": self.execute_tags_only.get(),
                    "llm_fallback_enabled": self.llm_fallback.get()
                }
            
            self.progress_var.set(30)
            self.status_var.set("正在分析文档...")
            self.log_message(f"输入文件: {self.input_file.get()}")
            
            # 构建最终的处理指令
            final_intent = ""
            
            # 如果启用了自动分页符，添加到指令前面
            if self.auto_page_break.get():
                final_intent += "首先在封面后插入分页符分隔封面和正文。"
                self.log_message("🔧 已启用自动分页符插入")
            
            # 如果启用了批注处理且有批注
            if self.use_comments.get() and self.extracted_comments:
                self.log_message(f"批注处理: 启用 ({len(self.extracted_comments)} 条批注)")
                
                # 构建批注指令
                comment_intent = "基于文档批注的处理指令:\n"
                for i, comment in enumerate(self.extracted_comments, 1):
                    scope = comment.get("scope_hint", "ANCHOR")
                    scope_desc = {"GLOBAL": "全文", "SECTION": "节级", "ANCHOR": "局部"}[scope]
                    comment_intent += f"{i}. [{scope_desc}范围] {comment['text']}\n"
                
                # 添加处理说明
                comment_intent += "\n请按照批注的作用域要求处理文档：\n"
                comment_intent += "- 全文范围：应用到整个文档\n"
                comment_intent += "- 节级范围：应用到相关章节\n"
                comment_intent += "- 局部范围：应用到批注标记的具体位置\n"
                
                if final_intent:
                    final_intent += "\n\n" + comment_intent
                else:
                    final_intent = comment_intent
                
                # 如果还有额外的用户指令，添加到后面
                if user_intent.strip():
                    final_intent += f"\n\n附加处理指令:\n{user_intent.strip()}\n"
                
                self.log_message(f"处理指令: 分页符 + 批注指令 + 附加指令")
            else:
                # 只有用户指令
                if user_intent.strip():
                    if final_intent:
                        final_intent += "\n\n" + user_intent.strip()
                    else:
                        final_intent = user_intent.strip()
                self.log_message(f"处理指令: {final_intent}")
            
            # 如果启用了封面保护，添加保护说明
            if self.cover_protection.get() and use_simple_pipeline:
                if final_intent:
                    final_intent += "\n\n重要：请保护封面和目录区域的格式，不要修改无页码区域的样式。"
                else:
                    final_intent = "请保护封面和目录区域的格式，不要修改无页码区域的样式。"
                self.log_message("🛡️ 已启用封面保护模式")
            
            # 如果没有任何指令，提供默认指令
            if not final_intent:
                final_intent = "请分析文档结构并进行基本的格式优化"
                self.log_message("使用默认处理指令")
            
            result = pipeline.process_document(self.input_file.get(), final_intent)
            
            self.progress_var.set(90)
            self.status_var.set("正在完成处理...")
            
            # 处理结果
            if hasattr(result, 'status') and result.status == "SUCCESS":
                self.progress_var.set(100)
                self.status_var.set("处理完成")
                self.log_message("✅ 文档处理成功！")
                
                # 获取输出文件路径
                output_path = getattr(result, 'output_path', None) or getattr(result, 'output_file', None)
                if output_path:
                    self.log_message(f"输出文件: {output_path}")
                
                # 获取审计目录
                audit_dir = getattr(result, 'audit_directory', None) or getattr(result, 'audit_dir', None)
                if audit_dir:
                    self.log_message(f"审计目录: {audit_dir}")
                
                # 如果有批注处理结果，显示详细信息
                if hasattr(result, 'comment_results') and result.comment_results:
                    self.log_message("📝 批注处理结果:")
                    for comment_result in result.comment_results:
                        status_icon = "✅" if comment_result.status == "APPLIED" else "⚠️"
                        self.log_message(f"  {status_icon} {comment_result.comment_id}: {comment_result.status}")
                
                # 显示封面保护状态
                if use_simple_pipeline and self.cover_protection.get():
                    self.log_message("🛡️ 封面保护功能已生效，封面和目录格式已保护")
                
                # 显示分页符插入状态
                if self.auto_page_break.get():
                    self.log_message("📄 已自动插入分页符分隔封面和正文")
                
                # 显示成功对话框
                success_msg = "文档处理完成！"
                if output_path:
                    success_msg += f"\n\n输出文件: {output_path}"
                if self.use_comments.get() and self.extracted_comments:
                    success_msg += f"\n\n处理了 {len(self.extracted_comments)} 条批注"
                if use_simple_pipeline and self.cover_protection.get():
                    success_msg += f"\n\n🛡️ 封面和目录格式已保护"
                if self.auto_page_break.get():
                    success_msg += f"\n\n📄 已插入分页符"
                
                self.root.after(0, lambda: messagebox.showinfo("成功", success_msg))
                
            else:
                # 处理失败或其他状态
                status = getattr(result, 'status', 'UNKNOWN')
                self.log_message(f"❌ 处理失败: {status}")
                
                error = getattr(result, 'error', None)
                if error:
                    self.log_message(f"错误信息: {error}")
                
                validation_errors = getattr(result, 'validation_errors', None)
                if validation_errors:
                    for error in validation_errors:
                        self.log_message(f"验证错误: {error}")
                
                # 显示错误对话框
                error_msg = f"处理失败: {status}"
                if error:
                    error_msg += f"\n\n错误: {error}"
                
                self.root.after(0, lambda: messagebox.showerror("处理失败", error_msg))
                
                self.progress_var.set(0)
                self.status_var.set("处理失败")
        
        except Exception as e:
            self.log_message(f"❌ 处理异常: {e}")
            self.root.after(0, lambda: messagebox.showerror("异常", f"处理过程中发生异常:\n\n{e}"))
            self.progress_var.set(0)
            self.status_var.set("处理异常")
        
        finally:
            # 恢复UI状态
            self.processing = False
            self.process_btn.config(state=NORMAL)
            self.stop_btn.config(state=DISABLED)
    
    def stop_processing(self):
        """停止处理"""
        self.processing = False
        self.progress_var.set(0)
        self.status_var.set("已停止")
        self.log_message("⏹️ 处理已停止")
        
        self.process_btn.config(state=NORMAL)
        self.stop_btn.config(state=DISABLED)
    
    def extract_comments_from_document(self, docx_path: str):
        """从Word文档中提取批注"""
        try:
            self.log_message(f"🔍 正在提取文档批注: {Path(docx_path).name}")
            
            import win32com.client
            
            # 打开Word应用
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            try:
                # 打开文档
                doc = word.Documents.Open(docx_path)
                
                # 提取批注
                comments = []
                for i, comment in enumerate(doc.Comments):
                    if not comment.Done:  # 只处理未解决的批注
                        comment_data = {
                            "comment_id": f"comment_{i+1}",
                            "author": comment.Author,
                            "created_time": str(comment.Date),
                            "text": comment.Range.Text.strip(),
                            "resolved": comment.Done,
                            "anchor": {
                                "paragraph_start": comment.Scope.Start,
                                "paragraph_end": comment.Scope.End,
                                "char_start": comment.Scope.Start,
                                "char_end": comment.Scope.End
                            },
                            "scope_hint": self.detect_scope_from_text(comment.Range.Text)
                        }
                        comments.append(comment_data)
                
                # 关闭文档
                doc.Close(False)
                
                # 保存批注数据
                self.extracted_comments = comments
                self.comments_json = {
                    "schema_version": "comments.v1",
                    "document_path": docx_path,
                    "extraction_time": time.strftime("%Y-%m-%d %H:%M:%S"),
                    "total_comments": len(comments),
                    "comments": comments
                }
                
                # 更新UI状态
                if comments:
                    self.comments_status.set(f"检测到 {len(comments)} 条批注")
                    self.dry_run_btn.config(state=NORMAL)
                    
                    # 自动保存批注JSON文件
                    self.auto_save_comments_json(docx_path)
                    
                    # 自动生成综合指令
                    combined_intent = self.generate_combined_intent_from_comments(comments)
                    if combined_intent:
                        self.user_intent.set(combined_intent)
                        self.intent_text.delete("1.0", END)
                        self.intent_text.insert("1.0", combined_intent)
                    
                    self.log_message(f"✅ 成功提取 {len(comments)} 条批注")
                    
                    # 显示批注预览
                    self.show_comments_preview(comments)
                else:
                    self.comments_status.set("未检测到批注")
                    self.dry_run_btn.config(state=DISABLED)
                    self.log_message("ℹ️ 文档中没有未解决的批注")
                
            finally:
                word.Quit()
                
        except Exception as e:
            self.log_message(f"❌ 批注提取失败: {e}")
            self.comments_status.set("批注提取失败")
            
    def detect_scope_from_text(self, comment_text: str) -> str:
        """从批注文本检测作用域"""
        text = comment_text.lower()
        
        # 检查显式标记
        if "scope=global" in text:
            return "GLOBAL"
        elif "scope=section" in text:
            return "SECTION"
        elif "scope=anchor" in text:
            return "ANCHOR"
        
        # 检查关键词
        global_keywords = ["全文", "全局", "全篇", "整体", "全文统一", "全文应用"]
        section_keywords = ["本节", "以该标题为范围", "以下段落"]
        
        if any(keyword in text for keyword in global_keywords):
            return "GLOBAL"
        elif any(keyword in text for keyword in section_keywords):
            return "SECTION"
        
        return "ANCHOR"  # 默认
    
    def generate_combined_intent_from_comments(self, comments: list) -> str:
        """从批注生成综合处理指令"""
        if not comments:
            return ""
        
        # 按作用域分组
        global_comments = [c for c in comments if c.get("scope_hint") == "GLOBAL"]
        section_comments = [c for c in comments if c.get("scope_hint") == "SECTION"]
        anchor_comments = [c for c in comments if c.get("scope_hint") == "ANCHOR"]
        
        intents = []
        
        # 添加全局指令
        for comment in global_comments:
            intents.append(f"全文范围: {comment['text']}")
        
        # 添加节级指令
        for comment in section_comments:
            intents.append(f"节级范围: {comment['text']}")
        
        # 添加锚点指令
        for comment in anchor_comments:
            intents.append(f"局部修改: {comment['text']}")
        
        if intents:
            combined = "基于文档批注的处理指令:\n" + "\n".join(f"- {intent}" for intent in intents)
            return combined
        
        return ""
    
    def show_comments_preview(self, comments: list):
        """显示批注预览"""
        preview_text = "\n📝 文档批注预览:\n" + "="*30 + "\n"
        
        for i, comment in enumerate(comments, 1):
            scope = comment.get("scope_hint", "ANCHOR")
            scope_desc = {"GLOBAL": "全文", "SECTION": "节级", "ANCHOR": "局部"}[scope]
            
            preview_text += f"{i}. [{scope_desc}] {comment['author']}: {comment['text'][:50]}...\n"
        
        preview_text += "="*30
        self.log_message(preview_text)
    
    def toggle_comment_processing(self):
        """切换批注处理模式"""
        if self.use_comments.get():
            self.log_message("✅ 已启用批注处理模式")
            if self.input_file.get():
                self.extract_comments_from_document(self.input_file.get())
        else:
            self.log_message("❌ 已禁用批注处理模式")
            self.comments_status.set("批注处理已禁用")
            self.dry_run_btn.config(state=DISABLED)
    
    def toggle_cover_protection(self):
        """切换封面保护模式"""
        if self.cover_protection.get():
            self.log_message("🛡️ 已启用封面保护功能")
            self.cover_status.set("封面保护: 启用")
            self.cover_status.config(foreground="green")
        else:
            self.log_message("⚠️ 已禁用封面保护功能")
            self.cover_status.set("封面保护: 禁用")
            self.cover_status.config(foreground="red")
    
    def preview_comments(self):
        """预览批注处理结果"""
        if not self.extracted_comments:
            messagebox.showwarning("警告", "没有检测到批注")
            return
        
        try:
            # 创建预览窗口
            preview_window = Toplevel(self.root)
            preview_window.title("批注处理预览")
            preview_window.geometry("800x600")
            
            # 创建文本区域
            text_area = scrolledtext.ScrolledText(preview_window, wrap=WORD)
            text_area.pack(fill=BOTH, expand=True, padx=10, pady=10)
            
            # 显示批注信息
            preview_content = "📝 批注处理预览\n" + "="*50 + "\n\n"
            
            for i, comment in enumerate(self.extracted_comments, 1):
                scope = comment.get("scope_hint", "ANCHOR")
                preview_content += f"批注 {i}:\n"
                preview_content += f"  作者: {comment['author']}\n"
                preview_content += f"  作用域: {scope}\n"
                preview_content += f"  内容: {comment['text']}\n"
                preview_content += f"  位置: 字符 {comment['anchor']['char_start']}-{comment['anchor']['char_end']}\n"
                preview_content += "-" * 40 + "\n\n"
            
            # 显示JSON结构
            preview_content += "\n📋 JSON结构预览:\n" + "="*50 + "\n"
            if self.comments_json:
                import json
                preview_content += json.dumps(self.comments_json, indent=2, ensure_ascii=False)
            
            text_area.insert("1.0", preview_content)
            text_area.config(state=DISABLED)
            
            # 添加保存按钮
            btn_frame = ttk.Frame(preview_window)
            btn_frame.pack(fill=X, padx=10, pady=(0, 10))
            
            ttk.Button(btn_frame, text="保存JSON", 
                      command=lambda: self.save_comments_json()).pack(side=LEFT, padx=(0, 10))
            ttk.Button(btn_frame, text="关闭", 
                      command=preview_window.destroy).pack(side=RIGHT)
            
        except Exception as e:
            messagebox.showerror("错误", f"预览失败: {e}")
    
    def auto_save_comments_json(self, docx_path: str):
        """自动保存批注JSON文件"""
        if not self.comments_json:
            return
        
        try:
            # 生成JSON文件名（与原文档同目录）
            docx_file = Path(docx_path)
            json_filename = docx_file.parent / f"{docx_file.stem}_comments.json"
            
            import json
            with open(json_filename, 'w', encoding='utf-8') as f:
                json.dump(self.comments_json, f, indent=2, ensure_ascii=False)
            
            self.log_message(f"💾 批注JSON已自动保存: {json_filename}")
            
            # 更新状态显示保存位置
            self.comments_status.set(f"检测到 {len(self.extracted_comments)} 条批注 (JSON已保存)")
            
        except Exception as e:
            self.log_message(f"⚠️ 自动保存批注JSON失败: {e}")
    
    def test_cover_protection(self):
        """测试封面保护功能"""
        if not self.input_file.get():
            messagebox.showwarning("警告", "请先选择输入文件")
            return
        
        if not self.api_key.get().strip():
            messagebox.showwarning("警告", "请先配置API密钥")
            return
        
        # 创建测试窗口
        test_window = Toplevel(self.root)
        test_window.title("封面保护功能测试")
        test_window.geometry("600x500")
        
        # 创建测试选项
        test_frame = ttk.LabelFrame(test_window, text="测试选项", padding="10")
        test_frame.pack(fill=X, padx=10, pady=10)
        
        # 测试类型选择
        test_type = StringVar(value="style_change")
        ttk.Radiobutton(test_frame, text="样式修改测试（1级标题改为楷体小四号2倍行距）", 
                       variable=test_type, value="style_change").pack(anchor=W, pady=2)
        ttk.Radiobutton(test_frame, text="分页符插入测试", 
                       variable=test_type, value="page_break").pack(anchor=W, pady=2)
        ttk.Radiobutton(test_frame, text="组合测试（分页符 + 样式修改）", 
                       variable=test_type, value="combined").pack(anchor=W, pady=2)
        
        # 测试结果显示
        result_frame = ttk.LabelFrame(test_window, text="测试结果", padding="10")
        result_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        result_text = scrolledtext.ScrolledText(result_frame, height=15, wrap=WORD)
        result_text.pack(fill=BOTH, expand=True)
        
        def run_test():
            """运行测试"""
            test_choice = test_type.get()
            result_text.delete("1.0", END)
            result_text.insert(END, "🧪 开始封面保护功能测试...\n")
            result_text.insert(END, "="*50 + "\n\n")
            
            # 根据测试类型生成测试指令
            if test_choice == "style_change":
                test_intent = "1级标题设置为楷体小四号2倍行距"
                result_text.insert(END, "测试类型: 样式修改测试\n")
                result_text.insert(END, "预期结果: 只修改正文中的1级标题，封面保持不变\n\n")
            elif test_choice == "page_break":
                test_intent = "插入分页符在封面后"
                result_text.insert(END, "测试类型: 分页符插入测试\n")
                result_text.insert(END, "预期结果: 在封面后插入分页符\n\n")
            else:  # combined
                test_intent = "插入分页符在封面后，1级标题设置为楷体小四号2倍行距"
                result_text.insert(END, "测试类型: 组合测试\n")
                result_text.insert(END, "预期结果: 插入分页符并修改正文样式，封面保持不变\n\n")
            
            result_text.insert(END, f"测试指令: {test_intent}\n")
            result_text.insert(END, f"封面保护: {'启用' if self.cover_protection.get() else '禁用'}\n")
            result_text.insert(END, "-"*50 + "\n\n")
            
            # 临时设置测试参数
            original_intent = self.user_intent.get()
            original_text = self.intent_text.get("1.0", END)
            
            try:
                # 设置测试指令
                self.user_intent.set(test_intent)
                self.intent_text.delete("1.0", END)
                self.intent_text.insert("1.0", test_intent)
                
                # 启动处理
                result_text.insert(END, "正在处理文档...\n")
                test_window.update()
                
                # 调用处理函数
                thread = threading.Thread(target=lambda: self.process_document(test_intent))
                thread.daemon = True
                thread.start()
                
                result_text.insert(END, "✅ 测试已启动，请查看主窗口的处理日志\n")
                result_text.insert(END, "处理完成后请检查输出文档验证封面保护效果\n")
                
            except Exception as e:
                result_text.insert(END, f"❌ 测试失败: {e}\n")
            finally:
                # 恢复原始设置
                self.user_intent.set(original_intent)
                self.intent_text.delete("1.0", END)
                self.intent_text.insert("1.0", original_text)
        
        # 按钮区域
        btn_frame = ttk.Frame(test_window)
        btn_frame.pack(fill=X, padx=10, pady=10)
        
        ttk.Button(btn_frame, text="🚀 开始测试", command=run_test).pack(side=LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="关闭", command=test_window.destroy).pack(side=RIGHT)
    
    def save_comments_json(self):
        """手动保存批注JSON文件"""
        if not self.comments_json:
            messagebox.showwarning("警告", "没有批注数据可保存")
            return
        
        filename = filedialog.asksaveasfilename(
            title="保存批注JSON",
            defaultextension=".json",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
        )
        
        if filename:
            try:
                import json
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(self.comments_json, f, indent=2, ensure_ascii=False)
                self.log_message(f"💾 批注JSON已保存: {filename}")
                messagebox.showinfo("成功", f"批注JSON已保存到:\n{filename}")
            except Exception as e:
                messagebox.showerror("错误", f"保存失败: {e}")
    
    def run_performance_test(self):
        """运行性能测试"""
        if messagebox.askyesno("性能测试", "是否运行性能测试？\n\n这将创建测试文档并执行多个测试用例。"):
            try:
                import subprocess
                subprocess.Popen([sys.executable, "性能测试工具.py"])
                self.log_message("🧪 性能测试已启动")
            except Exception as e:
                messagebox.showerror("错误", f"启动性能测试失败: {e}")
    
    def save_config(self):
        """保存配置"""
        config = {
            "llm": {
                "provider": self.api_provider.get(),
                "model": "gpt-4" if self.api_provider.get() == "openai" else "claude-3-sonnet-20240229",
                "api_key": self.api_key.get().strip(),
                "temperature": 0.1
            },
            "localization": {
                "language": "zh-CN",
                "style_aliases": {
                    "Heading 1": "标题 1",
                    "Heading 2": "标题 2",
                    "Normal": "正文"
                }
            },
            "validation": {
                "strict_mode": True,
                "rollback_on_failure": True
            }
        }
        
        try:
            with open("vnext_config.json", 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            
            self.log_message("💾 配置已保存")
            messagebox.showinfo("成功", "配置已保存到 vnext_config.json")
            
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败: {e}")
    
    def load_config(self):
        """加载配置"""
        try:
            # 首先尝试加载完整配置
            if Path("vnext_config.json").exists():
                with open("vnext_config.json", 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                # 加载OpenAI配置
                if "llm" in config and config["llm"].get("provider") == "openai":
                    self.openai_key = config["llm"].get("api_key", "")
                    self.api_provider.set("openai")
                    self.api_key.set(self.openai_key)
                elif "llm_backup" in config and config["llm_backup"].get("provider") == "openai":
                    self.openai_key = config["llm_backup"].get("api_key", "")
                
                # 加载Claude配置
                if "llm" in config and config["llm"].get("provider") == "anthropic":
                    self.claude_key = config["llm"].get("api_key", "")
                    self.api_provider.set("anthropic")
                    self.api_key.set(self.claude_key)
                elif "llm_backup" in config and config["llm_backup"].get("provider") == "anthropic":
                    self.claude_key = config["llm_backup"].get("api_key", "")
                
                # 如果主配置是OpenAI，备用是Claude
                if "llm" in config and config["llm"].get("provider") == "openai":
                    if "llm_backup" in config and config["llm_backup"].get("provider") == "anthropic":
                        self.claude_key = config["llm_backup"].get("api_key", "")
                
                # 如果主配置是Claude，备用是OpenAI
                elif "llm" in config and config["llm"].get("provider") == "anthropic":
                    if "llm_backup" in config and config["llm_backup"].get("provider") == "openai":
                        self.openai_key = config["llm_backup"].get("api_key", "")
                
                self.log_message("📂 完整配置已加载")
                
            # 如果没有完整配置，尝试简化配置
            elif Path("simple_config.json").exists():
                with open("simple_config.json", 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                # 加载两个密钥
                self.openai_key = config.get("openai_key", "")
                self.claude_key = config.get("claude_key", "")
                
                # 默认使用OpenAI
                self.api_provider.set("openai")
                self.api_key.set(self.openai_key)
                
                self.log_message("📂 简化配置已加载")
                self.log_message(f"🤖 OpenAI密钥: {'已配置' if self.openai_key else '未配置'}")
                self.log_message(f"🧠 Claude密钥: {'已配置' if self.claude_key else '未配置'}")
            
            else:
                self.log_message("⚠️ 未找到配置文件")
        
        except Exception as e:
            self.log_message(f"⚠️ 加载配置失败: {e}")
    
    def run(self):
        """运行GUI"""
        self.log_message("🚀 AutoWord vNext GUI 已启动")
        self.log_message("💡 提示: 请先配置API密钥，然后选择文档开始处理")
        self.root.mainloop()


def main():
    """主函数"""
    try:
        app = AutoWordGUI()
        app.run()
    except Exception as e:
        messagebox.showerror("启动错误", f"GUI启动失败:\n\n{e}")
        print(f"GUI启动失败: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()