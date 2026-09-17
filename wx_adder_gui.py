# -*- coding: utf-8 -*-
"""
微信自动添加好友工具 - GUI面板
前置条件：微信PC客户端(3.x)已登录并保持在前台。
"""
import threading
import queue
import random
import re
import time
import sys
import os
from pathlib import Path

import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

PHONE_RE = re.compile(r"1[3-9]\d{9}")


def normalizePhone(val):
    s = str(val).strip()
    s = re.sub(r"[\s\-]", "", s)
    if s.endswith(".0"):
        s = s[:-2]
    return s


def isValidPhone(val):
    if pd.isna(val):
        return False
    return bool(PHONE_RE.fullmatch(normalizePhone(val)))


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("微信自动添加好友工具")
        self.geometry("960x680")
        self.minsize(900, 640)

        self.excelPath = None
        self.df = None
        self.columns = []
        self.records = []
        self.validRecords = []

        self.taskRunning = False
        self.stopEvent = threading.Event()
        self.logQueue = queue.Queue()
        self.wx = None

        self._buildUI()
        self._pollLog()

        defaultExcel = Path("企业查询数据(1).xlsx")
        if defaultExcel.exists():
            self.loadExcel(str(defaultExcel))

    def _buildUI(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        topFrame = ctk.CTkFrame(self, corner_radius=8)
        topFrame.grid(row=0, column=0, padx=12, pady=(12, 6), sticky="ew")
        topFrame.grid_columnconfigure(1, weight=1)
        self._buildTop(topFrame)

        mainFrame = ctk.CTkFrame(self, corner_radius=8)
        mainFrame.grid(row=1, column=0, padx=12, pady=6, sticky="nsew")
        mainFrame.grid_columnconfigure(0, weight=1)
        mainFrame.grid_columnconfigure(1, weight=1)
        mainFrame.grid_rowconfigure(0, weight=1)
        self._buildLeft(mainFrame)
        self._buildRight(mainFrame)

        bottomFrame = ctk.CTkFrame(self, corner_radius=8, fg_color="transparent")
        bottomFrame.grid(row=2, column=0, padx=12, pady=(6, 12), sticky="ew")
        self._buildBottom(bottomFrame)

    def _buildTop(self, parent):
        ctk.CTkLabel(parent, text="微信自动添加好友工具", font=ctk.CTkFont(size=22, weight="bold")).grid(row=0, column=0, columnspan=5, padx=16, pady=(12, 8), sticky="w")

        ctk.CTkLabel(parent, text="Excel文件:").grid(row=1, column=0, padx=(16, 4), pady=4, sticky="w")
        self.fileEntry = ctk.CTkEntry(parent, state="readonly")
        self.fileEntry.grid(row=1, column=1, padx=4, pady=4, sticky="ew")
        ctk.CTkButton(parent, text="选择文件", width=90, command=self.selectExcel).grid(row=1, column=2, padx=4, pady=4)
        ctk.CTkButton(parent, text="刷新进度", width=90, fg_color="#6c757d", hover_color="#5a6268", command=self._updateInfo).grid(row=1, column=3, padx=4, pady=4)

        self.infoLabel = ctk.CTkLabel(parent, text="未加载文件", text_color="gray")
        self.infoLabel.grid(row=2, column=0, columnspan=5, padx=16, pady=(4, 10), sticky="w")

    def _buildLeft(self, parent):
        frame = ctk.CTkScrollableFrame(parent, corner_radius=8, label_text="参数配置")
        frame.grid(row=0, column=0, padx=(6, 3), pady=6, sticky="nsew")
        frame.grid_columnconfigure(1, weight=1)

        row = 0
        ctk.CTkLabel(frame, text="电话列:").grid(row=row, column=0, padx=8, pady=6, sticky="w")
        self.phoneColMenu = ctk.CTkOptionMenu(frame, values=["电话"], width=160)
        self.phoneColMenu.grid(row=row, column=1, padx=8, pady=6, sticky="ew")
        self.phoneColMenu.set("电话")

        row += 1
        ctk.CTkLabel(frame, text="备注名列:").grid(row=row, column=0, padx=8, pady=6, sticky="w")
        self.remarkColMenu = ctk.CTkOptionMenu(frame, values=["不设置备注"], width=160)
        self.remarkColMenu.grid(row=row, column=1, padx=8, pady=6, sticky="ew")
        self.remarkColMenu.set("企业名称" if "企业名称" in self.columns else "不设置备注")

        row += 1
        ctk.CTkLabel(frame, text="验证消息:").grid(row=row, column=0, padx=8, pady=6, sticky="nw")
        self.addMsgBox = ctk.CTkTextbox(frame, height=70)
        self.addMsgBox.grid(row=row, column=1, padx=8, pady=6, sticky="ew")
        self.addMsgBox.insert("1.0", "您好，我有业务想和您合作，希望能添加您为好友。")

        row += 1
        intervalFrame = ctk.CTkFrame(frame, fg_color="transparent")
        intervalFrame.grid(row=row, column=1, padx=8, pady=6, sticky="ew")
        intervalFrame.grid_columnconfigure((0, 2), weight=1)
        ctk.CTkLabel(frame, text="间隔(秒):").grid(row=row, column=0, padx=8, pady=6, sticky="w")
        self.minIntervalEntry = ctk.CTkEntry(intervalFrame, width=60)
        self.minIntervalEntry.insert(0, "8")
        self.minIntervalEntry.grid(row=0, column=0, padx=2, sticky="w")
        ctk.CTkLabel(intervalFrame, text="~").grid(row=0, column=1, padx=2)
        self.maxIntervalEntry = ctk.CTkEntry(intervalFrame, width=60)
        self.maxIntervalEntry.insert(0, "20")
        self.maxIntervalEntry.grid(row=0, column=2, padx=2, sticky="w")

        row += 1
        ctk.CTkLabel(frame, text="好友标签:").grid(row=row, column=0, padx=8, pady=6, sticky="w")
        self.tagsEntry = ctk.CTkEntry(frame, placeholder_text="多个用逗号分隔，如: 客户,中山")
        self.tagsEntry.grid(row=row, column=1, padx=8, pady=6, sticky="ew")

        row += 1
        ctk.CTkLabel(frame, text="最大数量:").grid(row=row, column=0, padx=8, pady=6, sticky="w")
        self.maxCountEntry = ctk.CTkEntry(frame, placeholder_text="0=不限制")
        self.maxCountEntry.insert(0, "0")
        self.maxCountEntry.grid(row=row, column=1, padx=8, pady=6, sticky="ew")

        row += 1
        ctk.CTkLabel(frame, text="起始行:").grid(row=row, column=0, padx=8, pady=6, sticky="w")
        self.startRowEntry = ctk.CTkEntry(frame, placeholder_text="0=从头开始，用于断点续传")
        self.startRowEntry.insert(0, "0")
        self.startRowEntry.grid(row=row, column=1, padx=8, pady=6, sticky="ew")

        row += 1
        ctk.CTkLabel(frame, text="已完成记录:").grid(row=row, column=0, padx=8, pady=6, sticky="w")
        self.progressFileEntry = ctk.CTkEntry(frame, placeholder_text="进度文件路径")
        self.progressFileEntry.insert(0, "progress.txt")
        self.progressFileEntry.grid(row=row, column=1, padx=8, pady=6, sticky="ew")

    def _buildRight(self, parent):
        frame = ctk.CTkFrame(parent, corner_radius=8)
        frame.grid(row=0, column=1, padx=(3, 6), pady=6, sticky="nsew")
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(frame, text="执行进度", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, padx=12, pady=(12, 4), sticky="w")

        statFrame = ctk.CTkFrame(frame, fg_color="transparent")
        statFrame.grid(row=1, column=0, padx=12, pady=4, sticky="ew")
        statFrame.grid_columnconfigure((0, 1, 2, 3), weight=1)
        self.statLabels = {}
        for i, key in enumerate(["total", "done", "success", "fail"]):
            colors = {"total": ("gray", "总数"), "done": ("blue", "已处理"), "success": ("green", "成功"), "fail": ("red", "失败")}
            color, label = colors[key]
            ctk.CTkLabel(statFrame, text=label, font=ctk.CTkFont(size=11), text_color="gray").grid(row=0, column=i, padx=4)
            v = ctk.CTkLabel(statFrame, text="0", font=ctk.CTkFont(size=20, weight="bold"), text_color=color)
            v.grid(row=1, column=i, padx=4)
            self.statLabels[key] = v

        self.progressBar = ctk.CTkProgressBar(frame, height=14)
        self.progressBar.grid(row=2, column=0, padx=12, pady=(8, 4), sticky="ew")
        self.progressBar.set(0)

        self.progressLabel = ctk.CTkLabel(frame, text="0 / 0", text_color="gray")
        self.progressLabel.grid(row=3, column=0, padx=12, pady=(0, 8), sticky="w")

        logFrame = ctk.CTkFrame(frame, fg_color="transparent")
        logFrame.grid(row=4, column=0, padx=12, pady=(4, 12), sticky="nsew")
        logFrame.grid_rowconfigure(0, weight=1)
        logFrame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(logFrame, text="实时日志", font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=0, sticky="nw", padx=0, pady=0)
        self.logBox = ctk.CTkTextbox(logFrame, height=220, state="disabled", wrap="none")
        self.logBox.grid(row=1, column=0, sticky="nsew", pady=(4, 0))

    def _buildBottom(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        self.wxStatusLabel = ctk.CTkLabel(parent, text="微信状态: 未连接", text_color="gray", font=ctk.CTkFont(size=13))
        self.wxStatusLabel.grid(row=0, column=0, padx=16, pady=10, sticky="w")

        self.startBtn = ctk.CTkButton(parent, text="开始添加", width=140, height=38, font=ctk.CTkFont(size=15, weight="bold"), command=self.startTask)
        self.startBtn.grid(row=0, column=1, padx=8, pady=10)

        self.stopBtn = ctk.CTkButton(parent, text="停止", width=100, height=38, fg_color="#d9534f", hover_color="#c9302c", state="disabled", command=self.stopTask)
        self.stopBtn.grid(row=0, column=2, padx=(8, 16), pady=10)

    def selectExcel(self):
        path = filedialog.askopenfilename(title="选择Excel文件", filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")])
        if path:
            self.loadExcel(path)

    def loadExcel(self, path):
        try:
            self.df = pd.read_excel(path, dtype=str)
        except Exception as e:
            messagebox.showerror("错误", f"读取Excel失败:\n{e}")
            return
        self.excelPath = path
        self.columns = list(self.df.columns)
        self.fileEntry.configure(state="normal")
        self.fileEntry.delete(0, "end")
        self.fileEntry.insert(0, path)
        self.fileEntry.configure(state="readonly")

        phoneCol = "电话" if "电话" in self.columns else self.columns[0]
        self.phoneColMenu.configure(values=self.columns)
        self.phoneColMenu.set(phoneCol)

        remarkCols = ["不设置备注", "法定代表人+电话"] + self.columns
        self.remarkColMenu.configure(values=remarkCols)
        defaultRemark = "法定代表人+电话" if "法定代表人" in self.columns else ("企业名称" if "企业名称" in self.columns else "不设置备注")
        self.remarkColMenu.set(defaultRemark)

        self._refreshRecords()
        self._updateInfo()
        self.log(f"已加载: {Path(path).name}，有效手机号 {len(self.validRecords)} 条")

    def _getAddedCount(self):
        pf = Path(self.progressFileEntry.get().strip() or "progress.txt")
        if pf.exists():
            return len([l for l in pf.read_text(encoding="utf-8").splitlines() if l.strip()])
        return 0

    def _updateInfo(self):
        valid = len(self.validRecords)
        invalid = len(self.records) - valid
        added = self._getAddedCount()
        pending = max(valid - added, 0)
        self.infoLabel.configure(text=f"总记录 {len(self.records)} | 有效 {valid} | 已添加 {added} | 待添加 {pending} | 无效 {invalid}", text_color="gray")

    def _refreshRecords(self):
        if self.df is None:
            return
        phoneCol = self.phoneColMenu.get()
        remarkCol = self.remarkColMenu.get()
        remarkCol = None if remarkCol == "不设置备注" else remarkCol
        self.records = []
        self.validRecords = []
        for idx, row in self.df.iterrows():
            phone = row[phoneCol] if phoneCol in self.df.columns else None
            remark = ""
            if remarkCol == "法定代表人+电话":
                parts = []
                for c in ["法定代表人", phoneCol]:
                    if c in self.df.columns:
                        v = row[c]
                        if not pd.isna(v):
                            s = str(v).strip()
                            if c == phoneCol:
                                s = normalizePhone(s)
                            parts.append(s)
                remark = "".join(parts)
            elif remarkCol and remarkCol in self.df.columns:
                v = row[remarkCol]
                if not pd.isna(v):
                    remark = str(v).strip()
            rec = {"row": idx + 2, "phone": phone, "remark": remark}
            self.records.append(rec)
            if isValidPhone(phone):
                self.validRecords.append(rec)

    def log(self, msg):
        ts = time.strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{ts}] {msg}"
        self.logQueue.put(line)
        try:
            with open("add_friends.log", "a", encoding="utf-8") as f:
                f.write(line + "\n")
        except Exception:
            pass

    def _pollLog(self):
        try:
            while True:
                msg = self.logQueue.get_nowait()
                self.logBox.configure(state="normal")
                self.logBox.insert("end", msg + "\n")
                self.logBox.see("end")
                self.logBox.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(200, self._pollLog)

    def updateStats(self, total, done, success, fail):
        self.statLabels["total"].configure(text=str(total))
        self.statLabels["done"].configure(text=str(done))
        self.statLabels["success"].configure(text=str(success))
        self.statLabels["fail"].configure(text=str(fail))
        if total > 0:
            self.progressBar.set(done / total)
            self.progressLabel.configure(text=f"{done} / {total}")

    def startTask(self):
        if self.df is None:
            messagebox.showwarning("提示", "请先选择Excel文件")
            return
        self._refreshRecords()
        if not self.validRecords:
            messagebox.showwarning("提示", "没有有效的手机号")
            return

        try:
            minInterval = float(self.minIntervalEntry.get())
            maxInterval = float(self.maxIntervalEntry.get())
            maxCount = int(self.maxCountEntry.get())
            startRow = int(self.startRowEntry.get())
        except ValueError:
            messagebox.showerror("错误", "间隔时间、最大数量、起始行必须是数字")
            return

        addMsg = self.addMsgBox.get("1.0", "end").strip()
        tagsRaw = self.tagsEntry.get().strip()
        tags = [t.strip() for t in tagsRaw.split(",") if t.strip()] if tagsRaw else None
        progressFile = self.progressFileEntry.get().strip() or "progress.txt"

        remarkCol = self.remarkColMenu.get()
        remarkCol = None if remarkCol == "不设置备注" else remarkCol

        params = {
            "addMsg": addMsg,
            "minInterval": minInterval,
            "maxInterval": maxInterval,
            "maxCount": maxCount,
            "startRow": startRow,
            "tags": tags,
            "remarkCol": remarkCol,
            "progressFile": progressFile,
        }

        self.taskRunning = True
        self.stopEvent.clear()
        self.startBtn.configure(state="disabled")
        self.stopBtn.configure(state="normal")
        self.wxStatusLabel.configure(text="微信状态: 连接中...", text_color="orange")

        t = threading.Thread(target=self.runTask, args=(params,), daemon=True)
        t.start()

    def stopTask(self):
        self.stopEvent.set()
        self.log("用户请求停止，等待当前操作完成...")

    def runTask(self, params):
        try:
            from wxauto import WeChat
        except ImportError:
            self.log("错误: wxauto未安装")
            self._taskEnd()
            return

        try:
            self.wx = WeChat()
            self.wxStatusLabel.configure(text="微信状态: 已连接", text_color="green")
            self.log("微信连接成功")
        except Exception as e:
            self.log(f"微信连接失败: {e}")
            self.log("请确保微信PC客户端(3.x)已登录并显示在前台")
            self.wxStatusLabel.configure(text="微信状态: 连接失败", text_color="red")
            self._taskEnd()
            return

        done = set()
        pf = Path(params["progressFile"])
        if pf.exists():
            done = set(l.strip() for l in pf.read_text(encoding="utf-8").splitlines() if l.strip())
        self.log(f"已完成记录 {len(done)} 个 (来自 {pf.name})")

        pending = [r for r in self.validRecords if normalizePhone(r["phone"]) not in done]
        if params["startRow"] > 0:
            pending = pending[params["startRow"]:]
            self.log(f"从第 {params['startRow']} 条开始")
        if params["maxCount"] > 0:
            pending = pending[:params["maxCount"]]
            self.log(f"限制最多 {params['maxCount']} 条")

        total = len(pending)
        self.log(f"本次待添加 {total} 条")
        if total == 0:
            self.log("没有待添加的记录")
            self._taskEnd()
            return

        successCount = 0
        failCount = 0
        results = []

        for i, rec in enumerate(pending, 1):
            if self.stopEvent.is_set():
                self.log("已停止")
                break

            phone = normalizePhone(rec["phone"])
            remark = rec["remark"] or None
            remarkStr = remark or "无备注"
            self.log(f"[{i}/{total}] 正在添加: {phone} ({remarkStr})")

            status = "unknown"
            errMsg = ""
            try:
                self.wx.AddNewFriend(
                    keywords=phone,
                    addmsg=params["addMsg"],
                    remark=remark,
                    tags=params["tags"],
                )
                status = "sent"
                successCount += 1
                with open(pf, "a", encoding="utf-8") as f:
                    f.write(phone + "\n")
                self.log(f"  [成功] {phone} | {remarkStr} | 已发送添加请求")
            except Exception as e:
                status = "fail"
                errMsg = str(e)
                failCount += 1
                self.log(f"  [失败] {phone} | {remarkStr} | {errMsg}")

            results.append({"行号": rec["row"], "手机号": phone, "备注": remark or "", "状态": status, "错误": errMsg, "时间": time.strftime("%Y-%m-%d %H:%M:%S")})
            self.updateStats(total, i, successCount, failCount)

            if i < total and not self.stopEvent.is_set():
                wait = random.uniform(params["minInterval"], params["maxInterval"])
                self.log(f"  等待 {wait:.1f}s")
                for _ in range(int(wait * 10)):
                    if self.stopEvent.is_set():
                        break
                    time.sleep(0.1)

        try:
            resultDf = pd.DataFrame(results)
            resultDf.to_excel("add_result.xlsx", index=False)
            self.log(f"结果已保存到 add_result.xlsx")
        except Exception as e:
            self.log(f"保存结果失败: {e}")

        self.log(f"完成! 成功 {successCount}，失败 {failCount}")
        self._taskEnd()

    def _taskEnd(self):
        self.taskRunning = False
        self.startBtn.configure(state="normal")
        self.stopBtn.configure(state="disabled")
        self._updateInfo()


if __name__ == "__main__":
    app = App()
    app.mainloop()
