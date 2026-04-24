#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
語音輸入工具 for Windows
按快捷鍵開始/停止錄音，自動轉錄並貼入目前視窗（不需要 ffmpeg）
支援 AI 後處理：移除填充詞、自動格式化、語句潤飾
"""

import os
import re
import ctypes
import threading
import time
from pathlib import Path

import numpy as np
import tkinter as tk
import pyaudio
import whisper
import keyboard
import pyperclip

# 讀取 .env（若有安裝 python-dotenv 則自動載入）
try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent / ".env")
except ImportError:
    pass


# ══════════════════════════════════════════════
#  設定區（可自行修改）
# ══════════════════════════════════════════════

HOTKEY     = "ctrl+shift+space"  # 全域快捷鍵
MODEL_SIZE = "small"             # base / small / medium / large-v3
LANGUAGE   = "zh"               # None=自動偵測  "zh"=強制中文  "en"=英文
SAMPLE_RATE = 16000
CHUNK       = 1024

# Whisper 提示：引導輸出繁體中文 + 混入英文 + 中文標點
WHISPER_PROMPT = "以下是繁體中文語音轉寫，含有少量英文詞彙，請使用繁體中文並加上完整標點符號。例：我今天去開會，討論 API 設計方案。"

# ── AI 後處理設定 ──────────────────────────────
# 將你的 Gemini API Key 寫在同目錄的 .env 檔案：
#   GEMINI_API_KEY=AIzaXXXX
# 取得免費 Key： https://aistudio.google.com/app/apikey
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
AI_MODEL       = "gemini-2.0-flash"   # 免費額度大，速度快
AI_DEFAULT_ON  = True                 # 啟動時預設開啟 AI 潤飾

AI_SYSTEM_PROMPT = """你是語音轉文字的後處理助手。請整理以下語音辨識的逐字稿：
1. 移除填充詞（嗯、呃、啊、喔、那個、就是、然後、對啊、所以說）
2. 移除明顯的重複或口誤（說了兩次同樣的話）
3. 自動加上適當的標點符號與分段
4. 若內容為列表或步驟，格式化為條列式（每項用「- 」開頭）
5. 潤飾語句使其通順自然，保留原意與語氣
直接輸出整理後的文字，不加任何說明、前綴或引號。"""

# ══════════════════════════════════════════════


class VoiceInput:
    """
    浮動語音輸入工具
    - 全域快捷鍵切換錄音
    - Whisper 本機轉錄（不需網路、不需 ffmpeg）
    - 自動貼入目前焦點視窗
    """

    def __init__(self):
        self.model = None
        self.is_recording = False
        self.audio_frames: list[bytes] = []
        self._pyaudio = pyaudio.PyAudio()
        self.ai_enabled = AI_DEFAULT_ON and bool(GEMINI_API_KEY)
        self._gemini_client = None
        if GEMINI_API_KEY:
            try:
                from google import genai
                self._gemini_client = genai.Client(api_key=GEMINI_API_KEY)
            except ImportError:
                self.ai_enabled = False

        self._build_ui()
        threading.Thread(target=self._load_model, daemon=True).start()
        keyboard.add_hotkey(HOTKEY, self._on_hotkey, suppress=True)

    # ── UI ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        self.root = tk.Tk()
        self.root.title("VoiceInput")
        self.root.overrideredirect(True)          # 去除標題列
        self.root.attributes("-topmost", True)    # 永遠在最上層
        self.root.attributes("-alpha", 0.0)       # 初始隱藏
        self.root.configure(bg="#1e1e2e")

        # 右下角定位
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        W, H = 350, 68
        self.root.geometry(f"{W}x{H}+{sw - W - 24}+{sh - H - 70}")

        # 狀態文字
        self.msg_var = tk.StringVar()
        self._lbl = tk.Label(
            self.root,
            textvariable=self.msg_var,
            bg="#1e1e2e",
            fg="#cdd6f4",
            font=("Microsoft JhengHei UI", 11),
            wraplength=330,
            justify="center",
        )
        self._lbl.pack(expand=True, fill="both", padx=10, pady=8)

        # 右鍵選單
        self._menu = tk.Menu(
            self.root, tearoff=0,
            bg="#313244", fg="#cdd6f4",
            activebackground="#45475a", activeforeground="#cdd6f4",
        )
        self._menu.add_command(label=self._ai_menu_label(), command=self._toggle_ai)
        self._menu.add_separator()
        self._menu.add_command(label="退出語音輸入", command=self._quit)
        self._lbl.bind("<Button-3>", lambda e: self._menu.post(e.x_root, e.y_root))

        # 100ms 後再套用 WS_EX_NOACTIVATE（視窗需先實際建立）
        self.root.after(100, self._apply_no_activate)

    def _apply_no_activate(self):
        """讓浮動視窗不搶奪其他視窗的焦點"""
        try:
            hwnd = int(self.root.wm_frame(), 16)
            GWL_EXSTYLE      = -20
            WS_EX_NOACTIVATE = 0x08000000
            WS_EX_TOOLWINDOW = 0x00000080
            style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            ctypes.windll.user32.SetWindowLongW(
                hwnd, GWL_EXSTYLE,
                style | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW,
            )
        except Exception:
            pass

    def _show(self, text: str, fg: str = "#cdd6f4", hide_after: int = 0):
        """從任何執行緒安全地更新 UI"""
        def _update():
            self.msg_var.set(text)
            self._lbl.config(fg=fg)
            self.root.attributes("-alpha", 0.93)
            if hide_after > 0:
                self.root.after(hide_after, self._hide)
        self.root.after(0, _update)

    def _hide(self):
        self.root.attributes("-alpha", 0.0)

    # ── 模型載入 ─────────────────────────────────────────────────────────

    def _ai_menu_label(self) -> str:
        if not GEMINI_API_KEY:
            return "AI 潤飾（需設定 GEMINI_API_KEY）"
        status = "✓ 開啟" if self.ai_enabled else "✗ 關閉"
        return f"AI 潤飾  [{status}]  點擊切換"

    def _toggle_ai(self):
        if not self._gemini_client:
            self._show("請先在 .env 設定 GEMINI_API_KEY", "#f38ba8", hide_after=4000)
            return
        self.ai_enabled = not self.ai_enabled
        self._menu.entryconfig(0, label=self._ai_menu_label())
        state = "開啟 ✨" if self.ai_enabled else "關閉"
        self._show(f"AI 潤飾已{state}", "#a6e3a1" if self.ai_enabled else "#6c7086", hide_after=2500)

    def _load_model(self):
        self._show("正在載入語音模型，請稍候...", "#fab387")
        try:
            self.model = whisper.load_model(MODEL_SIZE)
            ai_hint = "  ·  AI✨" if self.ai_enabled else ""
            self._show(
                f"就緒{ai_hint}  ·  {HOTKEY}  開始 / 停止錄音",
                "#a6e3a1",
                hide_after=5000,
            )
        except Exception as e:
            self._show(f"模型載入失敗：{e}", "#f38ba8")

    # ── 錄音控制 ─────────────────────────────────────────────────────────

    def _on_hotkey(self):
        if self.model is None:
            return
        if self.is_recording:
            self.is_recording = False   # 通知錄音執行緒停止
        else:
            self._start_recording()

    def _start_recording(self):
        self.is_recording = True
        self.audio_frames = []
        self._show("🔴  錄音中  ·  再按一次停止", "#f38ba8")
        threading.Thread(target=self._record_loop, daemon=True).start()

    def _record_loop(self):
        stream = self._pyaudio.open(
            format=pyaudio.paInt16,
            channels=1,
            rate=SAMPLE_RATE,
            input=True,
            frames_per_buffer=CHUNK,
        )
        while self.is_recording:
            try:
                data = stream.read(CHUNK, exception_on_overflow=False)
                self.audio_frames.append(data)
            except OSError:
                break
        stream.stop_stream()
        stream.close()
        threading.Thread(target=self._transcribe, daemon=True).start()

    # ── 轉錄 ─────────────────────────────────────────────────────────────

    def _transcribe(self):
        self._show("⏳  轉錄中...", "#89b4fa")

        if not self.audio_frames:
            self._show("（沒有錄到聲音）", "#6c7086", hide_after=3000)
            return

        try:
            # PCM int16 → float32，Whisper 直接接受 numpy array（不需 ffmpeg）
            raw = b"".join(self.audio_frames)
            audio_np = (
                np.frombuffer(raw, dtype=np.int16).astype(np.float32) / 32768.0
            )

            kwargs: dict = {"fp16": False, "task": "transcribe", "initial_prompt": WHISPER_PROMPT}
            if LANGUAGE:
                kwargs["language"] = LANGUAGE

            result = self.model.transcribe(audio_np, **kwargs)
            text: str = result["text"].strip()

        except Exception as e:
            self._show(f"轉錄失敗：{e}", "#f38ba8", hide_after=5000)
            return

        if not text:
            self._show("（沒有辨識到文字）", "#6c7086", hide_after=3000)
            return

        # 後處理：規則式清理（永遠執行）
        text = self._rule_clean(text)

        # AI 後處理（需要 Gemini API Key）
        if self.ai_enabled and self._gemini_client:
            text = self._ai_polish(text)

        # 複製到剪貼簿，貼入目前焦點視窗
        pyperclip.copy(text)
        time.sleep(0.12)
        keyboard.send("ctrl+v")

        preview = text[:46] + "…" if len(text) > 46 else text
        tag = "✨" if self.ai_enabled and self._gemini_client else "✓"
        self._show(f"{tag}  {preview}", "#a6e3a1", hide_after=4000)

    # ── 規則式後處理 ──────────────────────────────────────────────────────

    @staticmethod
    def _rule_clean(text: str) -> str:
        """移除填充詞、統一繁體標點、補齊句尾符號"""
        # ① 半形標點 → 全形繁體（中文字後面的分隔符號）
        punct_map = {"\uff0c": "，"}  # 全形連号 → 預留（已是全形）
        # 中文字後面的半形 , 轉。/ ，
        text = re.sub(r"(?<=[一-鿿]),", "，", text)
        text = re.sub(r"(?<=[一-鿿])\.", "。", text)
        text = re.sub(r"(?<=[一-鿿])!", "！", text)
        text = re.sub(r"(?<=[一-鿿])\?", "？", text)
        text = re.sub(r"(?<=[一-鿿]):", "：", text)
        text = re.sub(r"(?<=[一-鿿]);", "；", text)
        # 英文單字/數字後面緊接中文的 , 也轉
        text = re.sub(r"(?<=[a-zA-Z0-9]),(?=[一-鿿])", "，", text)
        # 尾隨空白和句末多餘的半形句點
        text = re.sub(r"[.,]。", "。", text)   # .。 或 ,。 去除前面多餘的

        # ② 移除填充詞（後面可選跟逗號/頓號）
        filler_patterns = [
            r"嗯+", r"呃+", r"啊+", r"喔+", r"哦+", r"欸+", r"唉+",
            r"那個嘛?", r"這個嘛?",
            r"就是說", r"所以說", r"你知道嗎?",
            r"然後呢?", r"接著呢?",
            r"對啊", r"對對", r"嗯嗯", r"好好",
        ]
        for p in filler_patterns:
            text = re.sub(rf"({p})[，、,]?", "", text)

        # ③ 清除句首多餘標點
        text = re.sub(r"^[，、,\s]+", "", text)

        # ④ 合併連續空白
        text = re.sub(r"[ \t]+", " ", text)

        # ⑤ 合併重複標點
        text = re.sub(r"([，。！？、])\1+", r"\1", text)
        text = re.sub(r"[，、]{2,}", "，", text)

        # ⑥ 標點前後多餘空格（保留英文單字前的空格）
        text = re.sub(r"\s*([，。！？、：；])\s*", r"\1", text)

        text = text.strip()

        # ⑦ 句尾補句號（英文句點結尾不補）
        if text and text[-1] not in "。！？…」』." and not text.endswith("..."):
            text += "。"

        return text

    # ── AI 後處理 ─────────────────────────────────────────────────────────

    def _ai_polish(self, raw_text: str) -> str:
        """呼叫 Gemini 移除填充詞、自動格式化、潤飾語句"""
        self._show("✨  AI 潤飾中...", "#cba6f7")
        try:
            from google import genai as _genai
            resp = self._gemini_client.models.generate_content(
                model=AI_MODEL,
                contents=raw_text,
                config=_genai.types.GenerateContentConfig(
                    system_instruction=AI_SYSTEM_PROMPT,
                    temperature=0.3,
                    max_output_tokens=2048,
                ),
            )
            return resp.text.strip()
        except Exception as e:
            self._show(f"AI 失敗，改用原始轉錄：{e}", "#f38ba8", hide_after=3000)
            return raw_text

    # ── 退出 ─────────────────────────────────────────────────────────────

    def _quit(self):
        keyboard.unhook_all()
        self._pyaudio.terminate()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    VoiceInput().run()
