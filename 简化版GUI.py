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
        
        # 用户意图区域
        intent_frame = ttk.LabelFrame(main_frame, text="处理指令", padding="10")
        intent_frame.grid(row=row, column=0, columnspan=3, sticky=(W, E), pady=(0, 10))
        intent_frame.columnconfigure(0, weight=1)
        row += 1
        
        # 预设指令
        preset_frame = ttk.Frame(intent_frame)
        preset_frame.grid(row=0, column=0, sticky=(W, E), pady=(0, 10))
        preset_frame.columnconfigure(0, weight=1)
        
        ttk.Label(preset_frame, text="常用指令:").grid(row=0, column=0, sticky=W)
        
        preset_buttons_frame = ttk.Frame(preset_frame)
        preset_buttons_frame.grid(row=1, column=0, sticky=(W, E), pady=(5, 0))
        
        presets = [
            ("删除摘要和参考文献", "删除摘要部分和参考文献部分，然后更新目录"),
            ("更新目录", "更新文档的目录"),
            ("标准化格式", "设置标题1为楷体12磅加粗，正文为宋体12磅，行距2倍"),
            ("删除摘要", "删除文档中的摘要部分")
        ]
        
        for i, (name, intent) in enumerate(presets):
            btn = ttk.Button(preset_buttons_frame, text=name, 
                           command=lambda i=intent: self.user_intent.set(i))
            btn.grid(row=i//2, column=i%2, sticky=W, padx=(0, 10), pady=2)
        
        # 自定义指令输入
        ttk.Label(intent_frame, text="自定义指令:").grid(row=1, column=0, sticky=W, pady=(10, 5))
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
        
        ttk.Button(control_frame, text="💾 保存配置", command=self.save_config).grid(row=0, column=2, padx=(0, 10))
        ttk.Button(control_frame, text="🧪 性能测试", command=self.run_performance_test).grid(row=0, column=3)
        
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
            
            # 导入AutoWord
            from autoword.vnext import VNextPipeline
            from autoword.vnext.core import VNextConfig, LLMConfig
            
            self.progress_var.set(20)
            self.status_var.set("正在配置AI模型...")
            
            # 创建配置
            llm_config = LLMConfig(
                provider=self.api_provider.get(),
                model="gpt-4" if self.api_provider.get() == "openai" else "claude-3-sonnet-20240229",
                api_key=self.api_key.get().strip(),
                temperature=0.1
            )
            
            config = VNextConfig(llm=llm_config)
            pipeline = VNextPipeline(config)
            
            self.progress_var.set(30)
            self.status_var.set("正在分析文档...")
            self.log_message(f"输入文件: {self.input_file.get()}")
            self.log_message(f"处理指令: {user_intent}")
            
            # 处理文档
            result = pipeline.process_document(self.input_file.get(), user_intent)
            
            self.progress_var.set(90)
            self.status_var.set("正在完成处理...")
            
            # 处理结果
            if result.status == "SUCCESS":
                self.progress_var.set(100)
                self.status_var.set("处理完成")
                self.log_message("✅ 文档处理成功！")
                self.log_message(f"输出文件: {result.output_path}")
                if result.audit_directory:
                    self.log_message(f"审计目录: {result.audit_directory}")
                
                # 显示成功对话框
                self.root.after(0, lambda: messagebox.showinfo(
                    "成功", 
                    f"文档处理完成！\n\n输出文件: {result.output_path}"
                ))
                
            else:
                self.log_message(f"❌ 处理失败: {result.status}")
                if result.error:
                    self.log_message(f"错误信息: {result.error}")
                if result.validation_errors:
                    for error in result.validation_errors:
                        self.log_message(f"验证错误: {error}")
                
                # 显示错误对话框
                error_msg = f"处理失败: {result.status}"
                if result.error:
                    error_msg += f"\n\n错误: {result.error}"
                
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
            if Path("vnext_config.json").exists():
                with open("vnext_config.json", 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                if "llm" in config:
                    llm = config["llm"]
                    self.api_provider.set(llm.get("provider", "openai"))
                    self.api_key.set(llm.get("api_key", ""))
                
                self.log_message("📂 配置已加载")
        
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