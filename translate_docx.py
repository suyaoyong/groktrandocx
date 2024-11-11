from openai import OpenAI
import os
from docx import Document
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.ttk import Progressbar
import time
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

class DocTranslator:
    def __init__(self):
        # 从环境变量获取 API key
        api_key = os.getenv('X_AI_API_KEY')
        if not api_key:
            raise ValueError("未找到 API key，请在 .env 文件中设置 X_AI_API_KEY")
            
        self.client = OpenAI(
            api_key=api_key,
            base_url="https://api.x.ai/v1"
        )
        
        # 支持的语言字典
        self.supported_languages = {
            "简体中文": "Chinese",
            "繁體中文": "Traditional Chinese",
            "日本語": "Japanese",
            "English": "English",
            "Español": "Spanish",
            "Français": "French",
            "Deutsch": "German",
            "한국어": "Korean",
            "Русский": "Russian",
            "Italiano": "Italian"
        }
        
    def translate_text(self, text, target_language):
        try:
            # 根据目标语言调整提示词
            if target_language == "English":
                prompt = f"Please translate the following text to English, maintaining professionalism and accuracy:\n\n{text}"
            elif target_language == "Japanese":
                prompt = f"以下のテキストを日本語に翻訳してください。専門性と正確性を保ちながら翻訳してください：\n\n{text}"
            else:
                prompt = f"请将以下文本翻译成{target_language}，保持专业性和准确性：\n\n{text}"

            completion = self.client.chat.completions.create(
                model="grok-beta",
                messages=[
                    {
                        "role": "system",
                        "content": f"You are a professional translator. Translate the text to {target_language} without adding any additional information or explanations."
                    },
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                temperature=0.3
            )
            return completion.choices[0].message.content
        except Exception as e:
            print(f"翻译出错: {str(e)}")
            return None

class TranslatorGUI:
    def __init__(self):
        try:
            self.translator = DocTranslator()
        except ValueError as e:
            messagebox.showerror("错误", str(e))
            raise
            
        self.window = tk.Tk()
        self.window.title("多语言文档翻译器")
        self.window.geometry("600x400")
        
        self.is_paused = False
        self.translation_start_time = None
        self.processed_paragraphs = 0
        
        # 创建界面元素
        self.setup_gui()
        
    def setup_gui(self):
        # 创建主框架
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 选择文件按钮
        self.select_button = ttk.Button(
            main_frame,
            text="选择Word文档",
            command=self.select_file,
            width=20
        )
        self.select_button.pack(pady=10)
        
        # 文件路径显示
        self.file_label = ttk.Label(main_frame, text="未选择文件", wraplength=500)
        self.file_label.pack(pady=5)
        
        # 语言选择框架
        lang_frame = ttk.LabelFrame(main_frame, text="翻译设置", padding="5")
        lang_frame.pack(fill=tk.X, pady=10, padx=5)
        
        # 目标语言选择
        ttk.Label(lang_frame, text="目标语言：").pack(side=tk.LEFT, padx=5)
        self.target_language = tk.StringVar(value="简体中文")
        self.language_combo = ttk.Combobox(
            lang_frame,
            textvariable=self.target_language,
            values=list(self.translator.supported_languages.keys()),
            state="readonly",
            width=20
        )
        self.language_combo.pack(side=tk.LEFT, padx=5)
        
        # 格式选项
        self.preserve_format = tk.BooleanVar(value=True)
        self.format_check = ttk.Checkbutton(
            lang_frame,
            text="保留原文档格式",
            variable=self.preserve_format
        )
        self.format_check.pack(side=tk.LEFT, padx=20)
        
        # 进度条
        self.progress = Progressbar(
            main_frame,
            orient=tk.HORIZONTAL,
            length=500,
            mode='determinate'
        )
        self.progress.pack(pady=10)
        
        # 状态标签
        self.status_label = ttk.Label(main_frame, text="")
        self.status_label.pack(pady=5)
        
        # 添加控制框架
        control_frame = ttk.LabelFrame(main_frame, text="翻译控制", padding="5")
        control_frame.pack(fill=tk.X, pady=10, padx=5)
        
        # 开始翻译按钮
        self.translate_button = ttk.Button(
            control_frame,
            text="开始翻译",
            command=self.start_translation,
            width=15,
            state=tk.DISABLED
        )
        self.translate_button.pack(side=tk.LEFT, padx=5)
        
        # 暂停/继续按钮
        self.pause_button = ttk.Button(
            control_frame,
            text="暂停",
            command=self.toggle_pause,
            width=15,
            state=tk.DISABLED
        )
        self.pause_button.pack(side=tk.LEFT, padx=5)
        
        # 清理缓存按钮
        self.clean_cache_button = ttk.Button(
            control_frame,
            text="清理缓存",
            command=self.clean_cache_with_confirm,
            width=15
        )
        self.clean_cache_button.pack(side=tk.LEFT, padx=5)
        
        # 状态信息框架
        status_frame = ttk.LabelFrame(main_frame, text="状态信息", padding="5")
        status_frame.pack(fill=tk.X, pady=10, padx=5)
        
        # 进度信息
        self.progress_label = ttk.Label(status_frame, text="进度: 0/0")
        self.progress_label.pack(side=tk.LEFT, padx=5)
        
        # 剩余时间
        self.time_label = ttk.Label(status_frame, text="预计剩余时间: --:--")
        self.time_label.pack(side=tk.LEFT, padx=5)
        
        # 缓存状态
        self.cache_label = ttk.Label(status_frame, text="缓存状态: 无")
        self.cache_label.pack(side=tk.LEFT, padx=5)
        
        # 添加说明文本
        help_text = "支持的语言���简体中文、繁體中文、日本語、English、Español、Français、Deutsch、한국어、Русский、Italiano"
        help_label = ttk.Label(main_frame, text=help_text, wraplength=500, foreground="gray")
        help_label.pack(pady=10)

    def toggle_pause(self):
        """切换暂停/继续状态"""
        self.is_paused = not self.is_paused
        if self.is_paused:
            self.pause_button.config(text="继续")
            self.status_label.config(text="翻译已暂停")
        else:
            self.pause_button.config(text="暂停")
            self.status_label.config(text="继续翻译...")

    def update_progress_info(self, current, total):
        """更新进度信息"""
        self.progress_label.config(text=f"进度: {current}/{total}")
        
        if self.translation_start_time and current > 0:
            elapsed_time = time.time() - self.translation_start_time
            avg_time_per_para = elapsed_time / current
            remaining_paras = total - current
            estimated_remaining_time = avg_time_per_para * remaining_paras
            
            # 转换为时分秒格式
            hours = int(estimated_remaining_time // 3600)
            minutes = int((estimated_remaining_time % 3600) // 60)
            seconds = int(estimated_remaining_time % 60)
            
            if hours > 0:
                time_str = f"{hours}小时{minutes}分钟"
            elif minutes > 0:
                time_str = f"{minutes}分钟{seconds}秒"
            else:
                time_str = f"{seconds}秒"
                
            self.time_label.config(text=f"预计剩余时间: {time_str}")

    def update_cache_status(self):
        """更新缓存状态"""
        try:
            cache_dir = os.path.join(os.path.dirname(self.file_path), ".translation_cache")
            if os.path.exists(cache_dir):
                cache_files = len(os.listdir(cache_dir))
                cache_size = sum(os.path.getsize(os.path.join(cache_dir, f)) 
                               for f in os.listdir(cache_dir))
                size_str = f"{cache_size/1024/1024:.1f}MB" if cache_size > 1024*1024 else f"{cache_size/1024:.1f}KB"
                self.cache_label.config(text=f"缓存状态: {cache_files}个文件 ({size_str})")
            else:
                self.cache_label.config(text="缓存状态: 无")
        except:
            self.cache_label.config(text="缓存状态: 未知")

    def clean_cache_with_confirm(self):
        """清理缓存并确认"""
        if messagebox.askyesno("确认", "确定要清理所有缓存文件吗？"):
            self.clean_cache()
            self.update_cache_status()
            messagebox.showinfo("成功", "缓存已清理")

    def start_translation(self):
        if not hasattr(self, 'file_path'):
            messagebox.showerror("错误", "请先选择文件")
            return
            
        try:
            # 重置计时器和进度
            self.translation_start_time = time.time()
            self.processed_paragraphs = 0
            
            # 启用暂停按钮
            self.pause_button.config(state=tk.NORMAL)
            
            # 获取选择的目标语言
            target_language = self.translator.supported_languages[self.target_language.get()]
            
            # 创建缓存文件路径
            cache_dir = os.path.join(os.path.dirname(self.file_path), ".translation_cache")
            if not os.path.exists(cache_dir):
                os.makedirs(cache_dir)
                
            cache_file = os.path.join(
                cache_dir, 
                f"{os.path.basename(self.file_path)}_{target_language}_cache.docx"
            )
            progress_file = os.path.join(
                cache_dir,
                f"{os.path.basename(self.file_path)}_{target_language}_progress.txt"
            )
            
            # 检查是否存在未完成的翻译
            last_index = 0
            doc = Document(self.file_path)
            new_doc = None
            
            if os.path.exists(cache_file) and os.path.exists(progress_file):
                with open(progress_file, 'r') as f:
                    last_index = int(f.read().strip() or '0')
                    
                if last_index > 0:
                    response = messagebox.askyesno(
                        "发现未完成的翻译",
                        f"发现上次翻译到第 {last_index} 段，是否继续上次的翻译？"
                    )
                    if response:
                        new_doc = Document(cache_file)
                    else:
                        last_index = 0
            
            if new_doc is None:
                new_doc = Document()
                
            total_paragraphs = len(doc.paragraphs)
            self.progress['maximum'] = total_paragraphs
            
            # 从上次的位置继续翻译
            for i, para in enumerate(doc.paragraphs[last_index:], start=last_index):
                while self.is_paused:
                    self.window.update()
                    time.sleep(0.1)
                    continue
                    
                try:
                    self.processed_paragraphs = i + 1
                    self.update_progress_info(i + 1, total_paragraphs)
                    self.status_label.config(text=f"正在翻译第 {i+1}/{total_paragraphs} 段...")
                    self.progress['value'] = i + 1
                    self.window.update()
                    
                    if para.text.strip():
                        new_para = new_doc.add_paragraph()
                        if self.preserve_format.get():
                            try:
                                new_para.style = para.style
                            except:
                                pass
                        
                        translated_text = self.translator.translate_text(para.text, target_language)
                        if translated_text:
                            new_para.text = translated_text
                        else:
                            new_para.text = para.text
                    else:
                        new_doc.add_paragraph()
                    
                    # 每翻译完一段就保存进度
                    new_doc.save(cache_file)
                    with open(progress_file, 'w') as f:
                        f.write(str(i + 1))
                    
                    time.sleep(0.5)
                    
                    # 更新缓存状态
                    self.update_cache_status()
                    
                except Exception as para_error:
                    print(f"翻译段落 {i+1} 时出错: {str(para_error)}")
                    # 保存当前进度并提示用户
                    new_doc.save(cache_file)
                    messagebox.showwarning(
                        "警告",
                        f"翻译第 {i+1} 段时出现错误，已保存进度。\n您可以稍后继续翻译。"
                    )
                    return
            
            # 翻译完成，保存最终文档
            output_path = os.path.splitext(self.file_path)[0] + f"_translated_{target_language}.docx"
            new_doc.save(output_path)
            
            # 清理缓存文件
            try:
                os.remove(cache_file)
                os.remove(progress_file)
            except:
                pass
            
            self.status_label.config(text="翻译完成！")
            messagebox.showinfo("成功", f"翻译已完成！\n保存至: {output_path}")
            
        except Exception as e:
            messagebox.showerror("错误", f"翻译过程中出错：{str(e)}")
            self.status_label.config(text="翻译失败")
        
        finally:
            self.pause_button.config(state=tk.DISABLED)
            self.progress['value'] = 0
            self.update_cache_status()

    def select_file(self):
        """选择要翻译的Word文档"""
        file_path = filedialog.askopenfilename(
            title="选择Word文档",
            filetypes=[("Word文档", "*.docx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.file_path = file_path
            self.file_label.config(text=f"已选择: {file_path}")
            self.translate_button.config(state=tk.NORMAL)

    def diagnose_document(self, doc_path):
        """诊断文档问题"""
        try:
            # 检查文件大小
            file_size = os.path.getsize(doc_path)
            print(f"文件大小: {file_size} 字节")

            # 尝试以二进制方式读取
            with open(doc_path, 'rb') as f:
                content = f.read()
                # 检查文件头部是否为有效的 Office Open XML 格式
                if not content.startswith(b'PK'):
                    return "文件不是有效的 Office Open XML 格式"

            # 尝试打开文档并分析内容
            doc = Document(doc_path)
            
            # 收集文档信息
            info = {
                "段落数": len(doc.paragraphs),
                "样式数": len(doc.styles),
                "节数": len(doc.sections),
                "表格数": len(doc.tables)
            }
            
            # 检查段落内容
            problematic_paras = []
            for i, para in enumerate(doc.paragraphs):
                try:
                    # 尝试访问段落属性
                    style = para.style
                    format = para.paragraph_format
                    runs = para.runs
                except Exception as e:
                    problematic_paras.append(f"段落 {i+1}: {str(e)}")

            return {
                "基本信息": info,
                "问题段落": problematic_paras
            }

        except Exception as e:
            return f"诊断时发生错误: {str(e)}"

    def run(self):
        """运行应用程序"""
        self.window.mainloop()

    def clean_cache(self):
        """清理所有缓存文件"""
        try:
            cache_dir = os.path.join(os.path.dirname(self.file_path), ".translation_cache")
            if os.path.exists(cache_dir):
                for file in os.listdir(cache_dir):
                    try:
                        os.remove(os.path.join(cache_dir, file))
                    except:
                        pass
                os.rmdir(cache_dir)
        except:
            pass

def main():
    app = TranslatorGUI()
    app.run()

if __name__ == "__main__":
    main() 