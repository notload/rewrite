import pandas as pd
import random
import os
from pathlib import Path
from datetime import datetime
import numpy as np
import requests
import tempfile
import time
import subprocess
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.clock import Clock
from kivy.utils import platform
from kivy.core.audio import SoundLoader   # 用于播放音频（安卓支持）

# 在安卓上，使用 SoundLoader 替代 playsound
if platform == 'android':
    import android  # 用于请求权限等

class DictationApp(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        
        # 输出显示区域（滚动）
        self.output = Label(text="欢迎使用听写程序\n请先加载 Excel 文件", 
                            size_hint_y=None, height=400, 
                            text_size=(400, None), halign='left', valign='top')
        scroll = ScrollView(size_hint=(1, 0.8))
        scroll.add_widget(self.output)
        self.add_widget(scroll)
        
        # 输入区域
        input_box = BoxLayout(size_hint=(1, 0.1))
        self.input = TextInput(hint_text="输入答案或指令", multiline=False)
        send_btn = Button(text="发送", size_hint_x=0.2)
        send_btn.bind(on_press=self.on_send)
        input_box.add_widget(self.input)
        input_box.add_widget(send_btn)
        self.add_widget(input_box)
        
        # 状态栏
        self.status = Label(text="状态：等待加载", size_hint=(1, 0.1))
        self.add_widget(self.status)
        
        # 初始化全局变量（与您的原脚本一致）
        self.df = None
        self.mode = None
        self.col_pick = None
        self.col_show = None
        self.col_update = None
        self.extract_counts = {}
        self.answer_counts = {}
        self.total_questions = 0
        self.correct_answers = 0
        self.idx = None          # 当前抽取的行索引
        self.row = None
        self.waiting_for_answer = False   # 是否等待用户输入答案
        self.user_answer = None
        
        # 文件路径（由用户通过输入框指令设定）
        self.file_path = None
        
        self.append_output("请先输入 'load 文件路径' 来加载 Excel")
        self.append_output("示例：load /sdcard/听写文件.xlsx")
    
    def append_output(self, text):
        """向输出区域追加文本"""
        current = self.output.text
        self.output.text = current + "\n" + text
        # 自动滚动到底部（可优化）
    
    def on_send(self, instance):
        """处理用户输入"""
        cmd = self.input.text.strip()
        self.input.text = ""
        if not cmd:
            return
        self.process_command(cmd)
    
    def process_command(self, cmd):
        """解析用户命令"""
        if cmd.lower().startswith("load "):
            path = cmd[5:].strip()
            self.load_file(path)
        elif cmd.lower() == "stop":
            self.save_and_exit()
        elif self.waiting_for_answer:
            # 当前等待答案输入
            self.user_answer = cmd
            self.waiting_for_answer = False
            self.continue_question()
        else:
            self.append_output("未知指令，请使用 'load 文件路径' 加载文件，或输入答案")
    
    def load_file(self, path):
        """加载 Excel 并选择模式"""
        try:
            self.df = pd.read_excel(path, header=None)
            self.append_output(f"成功加载文件：{path}，共 {len(self.df)} 行")
            self.file_path = path
            # 选择模式（简化：用户输入模式）
            self.append_output("请选择模式：输入 'mode 1' 或 'mode 2'")
            self.status.text = "请选择模式"
        except Exception as e:
            self.append_output(f"加载失败：{e}")
    
    def set_mode(self, mode):
        if mode == 1:
            self.col_pick, self.col_show, self.col_update = 0, 1, 2
        elif mode == 2:
            self.col_pick, self.col_show, self.col_update = 1, 0, 3
        else:
            self.append_output("无效模式，请输入 1 或 2")
            return
        self.mode = mode
        self.append_output(f"模式已设为 {mode}")
        self.status.text = "开始听写..."
        # 启动第一次抽取
        self.next_question()
    
    def next_question(self):
        """抽取下一个词条"""
        if self.df is None:
            self.append_output("请先加载文件")
            return
        
        # 加权选择
        idx = self.weighted_choice()
        self.idx = idx
        self.row = self.df.loc[idx]
        
        # 更新抽取次数
        self.extract_counts[idx] = self.extract_counts.get(idx, 0) + 1
        extract_count = self.extract_counts[idx]
        error_count = int(self.row[self.col_update])
        answer_count = self.answer_counts.get(idx, 0)
        error_rate = (error_count / answer_count * 100) if answer_count > 0 else 0.0
        
        # 显示提示
        self.append_output(f"\n--- 第 {self.total_questions+1} 题 ---")
        self.append_output(f"词条：{self.row[self.col_pick]}")
        self.append_output(f"总抽取：{extract_count}，已作答：{answer_count}，错误：{error_count}，错误率：{error_rate:.1f}%")
        if error_count >= 5:
            self.append_output("⚠️ 错误较多，重点复习！")
        elif error_count >= 3:
            self.append_output("⚡ 错误较多，认真作答！")
        
        self.append_output("请输入你的答案（或输入 'skip' 查看答案）")
        self.waiting_for_answer = True
        self.status.text = "等待输入答案"
    
    def weighted_choice(self):
        error_counts = self.df[self.col_update].values.astype(float)
        weights = (error_counts + 1) ** 2
        if np.all(weights == weights[0]):
            return random.choice(self.df.index)
        return random.choices(self.df.index, weights=weights, k=1)[0]
    
    def continue_question(self):
        """用户已输入答案，进行处理"""
        user_answer = self.user_answer.strip()
        self.user_answer = None
        
        if user_answer.lower() == "skip":
            self.append_output(f"正确内容：{self.row[self.col_show]}")
            # 播放读音（可选项）
            self.append_output("如需读音，输入 'play'")
            self.waiting_for_answer = True  # 等待 play 指令
            return
        
        # 自动匹配检测
        correct = str(self.row[self.col_show]).strip().lower()
        if user_answer and user_answer.lower() == correct:
            self.append_output("✅ 自动匹配正确！")
            answer = 'y'
        else:
            # 手动确认
            self.append_output(f"你的答案：{user_answer}，正确内容：{self.row[self.col_show]}")
            self.append_output("输入 'y' 表示正确，'n' 表示错误")
            self.waiting_for_answer = True
            # 保存当前状态，等待 y/n
            self.answer_pending = True
            return
        
        # 处理正确/错误
        self.handle_judgement(answer)
    
    def handle_judgement(self, answer):
        # 更新作答次数
        idx = self.idx
        self.answer_counts[idx] = self.answer_counts.get(idx, 0) + 1
        
        if answer.lower() == 'y':
            self.correct_answers += 1
            self.append_output("✅ 回答正确！")
        else:
            # 错误，记录错误次数
            self.df.at[idx, self.col_update] += 1
            new_error = int(self.df.at[idx, self.col_update])
            self.append_output(f"❌ 错误已记录，当前错误次数：{new_error}")
            # 询问是否播放读音
            self.append_output("播放读音？输入 'play' 或 'no'")
            self.waiting_for_answer = True
            # 设置状态，等待 play/no
            self.audio_pending = True
            return
        
        self.total_questions += 1
        self.status.text = f"已完成 {self.total_questions} 题"
        # 继续下一题
        Clock.schedule_once(lambda dt: self.next_question(), 0.1)
    
    # 处理 play 等指令（已在 process_command 中捕捉）
    def process_command(self, cmd):
        if cmd.lower().startswith("load "):
            path = cmd[5:].strip()
            self.load_file(path)
        elif cmd.lower() == "stop":
            self.save_and_exit()
        elif cmd.lower().startswith("mode "):
            try:
                mode = int(cmd[5:])
                self.set_mode(mode)
            except:
                self.append_output("格式错误，输入 'mode 1' 或 'mode 2'")
        elif cmd.lower() == "play":
            # 播放当前词的读音
            if self.row is not None:
                self.play_audio(str(self.row[self.col_show]))
            else:
                self.append_output("无内容可播放")
            # 如果是在等待 play 状态，则继续
            if hasattr(self, 'audio_pending') and self.audio_pending:
                self.audio_pending = False
                # 继续下一题（因为播放后可继续）
                self.total_questions += 1
                self.status.text = f"已完成 {self.total_questions} 题"
                Clock.schedule_once(lambda dt: self.next_question(), 0.1)
            elif hasattr(self, 'answer_pending') and self.answer_pending:
                # 这是在手动确认时等待 y/n 的情况
                pass
        elif cmd.lower() in ('y', 'n'):
            # 处理手动确认
            if hasattr(self, 'answer_pending') and self.answer_pending:
                self.answer_pending = False
                self.handle_judgement(cmd)
            else:
                self.append_output("当前没有等待确认")
        elif cmd.lower() == "skip":
            # 已在其他地方处理
            pass
        else:
            # 如果处于等待答案状态，直接当作答案
            if self.waiting_for_answer:
                self.user_answer = cmd
                self.waiting_for_answer = False
                self.continue_question()
            else:
                self.append_output("未知指令")
    
    def play_audio(self, text):
        # 百度语音接口
        lang = "zh" if any('\u4e00' <= ch <= '\u9fff' for ch in text) else "en"
        url = "http://fanyi.baidu.com/gettts"
        params = {"lan": lang, "text": text, "spd": 3, "source": "web"}
        try:
            resp = requests.get(url, params=params, timeout=5)
            if resp.status_code == 200:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".mp3") as f:
                    f.write(resp.content)
                    tmp_path = f.name
                if platform == 'android':
                    # 使用 Kivy 的 SoundLoader
                    sound = SoundLoader.load(tmp_path)
                    if sound:
                        sound.play()
                        # 等待播放结束（简单延时）
                        time.sleep(sound.length + 0.5)
                    else:
                        self.append_output("播放失败")
                else:
                    # 电脑上使用子进程
                    subprocess.run(["ffplay", "-nodisp", "-autoexit", tmp_path], check=False)
                os.unlink(tmp_path)
            else:
                self.append_output("获取语音失败")
        except Exception as e:
            self.append_output(f"播放出错：{e}")
    
    def save_and_exit(self):
        # 保存文件（与原逻辑类似）
        if self.df is None:
            self.append_output("没有数据可保存")
            return
        # 这里简化保存，可添加保存路径选择
        now = datetime.now()
        month = f"{now.month:02d}"
        day = f"{now.day:02d}"
        prefix = f"{month}_{day}"
        save_dir = Path("/sdcard/听写记录")
        save_dir.mkdir(parents=True, exist_ok=True)
        # 寻找最大编号
        max_num = 0
        for f in save_dir.glob(f"{prefix}_*.xlsx"):
            stem = f.stem
            parts = stem.split('_')
            if len(parts) == 3 and parts[2].isdigit():
                num = int(parts[2])
                if num > max_num:
                    max_num = num
        next_num = max_num + 1
        filename = f"{prefix}_{next_num}.xlsx"
        save_path = save_dir / filename
        # 保存
        try:
            self.df.to_excel(save_path, index=False, header=False)
            self.append_output(f"✅ 已保存至：{save_path}")
        except Exception as e:
            self.append_output(f"保存失败：{e}")
        # 退出应用
        App.get_running_app().stop()

class DictationAppMain(App):
    def build(self):
        return DictationApp()

if __name__ == '__main__':
    DictationAppMain().run()