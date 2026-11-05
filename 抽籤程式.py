# -*- coding: utf-8 -*-
"""
班級工作與區域分配抽籤程式
基於「餐桶分配表.docx」與「115-1公共區域分配表.docx」開發
支援：
1. 快速切換檔案內建預設範本
2. 動態新增、刪除小項目
3. 自訂各項目名額人數
4. 自行輸入抽籤名單 (支援姓名、座號、換行或逗號)
5. 公平隨機抽籤與結果匯出 (文字/CSV)
"""

import sys
import os
import random
import csv
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# 內建預設範本資料
PRESETS = {
    "餐桶分配表 (15人)": {
        "title": "抬餐桶工作分配表",
        "items": [
            ("抬餐桶", "湯桶", "", 2),
            ("抬餐桶", "飯桶", "", 2),
            ("抬餐桶", "菜桶(主)", "", 2),
            ("抬餐桶", "菜桶(副)", "", 2),
            ("抬餐桶", "低年級支援", "", 4),
            ("抬餐桶", "湯匙", "", 1),
            ("抬餐桶", "水果", "", 1),
            ("抬餐桶", "預備班", "", 1),
        ]
    },
    "僅公共打掃區域 (15人)": {
        "title": "公共打掃區域分配表",
        "items": [
            ("公共打掃", "一樓廁所", "", 3),
            ("公共打掃", "車庫", "含車庫、洗滌槽、車道、校門口", 8),
            ("公共打掃", "東樓梯", "一樓至四樓含教室走廊", 3),
            ("公共打掃", "走廊", "", 1),
        ]
    },
    "僅班上打掃區域 (15人)": {
        "title": "班上打掃區域分配表",
        "items": [
            ("班上打掃", "掃地、拖地", "", 4),
            ("班上打掃", "窗戶", "", 3),
            ("班上打掃", "桌櫃", "", 4),
            ("班上打掃", "黑板、餐桌", "", 1),
            ("班上打掃", "走廊", "", 3),
        ]
    }
}

class LotteryApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("🎯 班級工作與區域分配抽籤系統")
        self.geometry("980x720")
        self.minsize(850, 600)

        # 設定風格
        self.style = ttk.Style(self)
        if "vista" in self.style.theme_names():
            self.style.theme_use("vista")
        elif "clam" in self.style.theme_names():
            self.style.theme_use("clam")

        self.last_results = []
        self.last_unassigned = []

        self._create_widgets()
        self.load_preset("餐桶分配表 (15人)")

    def _create_widgets(self):
        # 頂部範本列
        top_frame = ttk.LabelFrame(self, text="📋 快速載入檔案範本", padding=10)
        top_frame.pack(fill=tk.X, padx=12, pady=(10, 5))

        self.preset_var = tk.StringVar(value="餐桶分配表 (15人)")
        for key in PRESETS.keys():
            r = ttk.Radiobutton(top_frame, text=key, value=key, variable=self.preset_var, command=lambda k=key: self.load_preset(k))
            r.pack(side=tk.LEFT, padx=8)

        # 主要內容分割區 (左: 項目與名額, 右: 抽籤名單)
        main_paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=12, pady=5)

        # === 左側：分配項目管理 ===
        left_frame = ttk.LabelFrame(main_paned, text="📝 工作分配項目與名額", padding=10)
        main_paned.add(left_frame, weight=3)

        # 項目表格 (Treeview)
        tree_frame = ttk.Frame(left_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(tree_frame, columns=("category", "name", "desc", "quota"), show="headings", selectmode="browse")
        self.tree.heading("category", text="分類/組別")
        self.tree.heading("name", text="項目名稱")
        self.tree.heading("desc", text="說明/備註")
        self.tree.heading("quota", text="名額(人數)")

        self.tree.column("category", width=80, anchor=tk.CENTER)
        self.tree.column("name", width=110, anchor=tk.W)
        self.tree.column("desc", width=140, anchor=tk.W)
        self.tree.column("quota", width=70, anchor=tk.CENTER)

        tree_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 項目總計資訊
        self.quota_summary_var = tk.StringVar(value="總計：0 項，總需求人數：0 人")
        lbl_summary = ttk.Label(left_frame, textvariable=self.quota_summary_var, font=("Microsoft JhengHei", 9, "bold"), foreground="#2563eb")
        lbl_summary.pack(anchor=tk.W, pady=(6, 4))

        # 新增項目區塊
        add_frame = ttk.Frame(left_frame)
        add_frame.pack(fill=tk.X, pady=(4, 0))

        ttk.Label(add_frame, text="組別:").grid(row=0, column=0, padx=2)
        self.entry_cat = ttk.Entry(add_frame, width=8)
        self.entry_cat.grid(row=0, column=1, padx=2)

        ttk.Label(add_frame, text="名稱:").grid(row=0, column=2, padx=2)
        self.entry_name = ttk.Entry(add_frame, width=12)
        self.entry_name.grid(row=0, column=3, padx=2)

        ttk.Label(add_frame, text="備註:").grid(row=0, column=4, padx=2)
        self.entry_desc = ttk.Entry(add_frame, width=12)
        self.entry_desc.grid(row=0, column=5, padx=2)

        ttk.Label(add_frame, text="名額:").grid(row=0, column=6, padx=2)
        self.spin_quota = ttk.Spinbox(add_frame, from_=1, to=99, width=4)
        self.spin_quota.set(1)
        self.spin_quota.grid(row=0, column=7, padx=2)

        btn_add = ttk.Button(add_frame, text="➕ 新增項目", command=self.add_item)
        btn_add.grid(row=0, column=8, padx=(6, 2))

        # 項目操作按鈕列 (名額+1、名額-1、刪除項目)
        act_frame = ttk.Frame(left_frame)
        act_frame.pack(fill=tk.X, pady=(6, 0))

        btn_plus = ttk.Button(act_frame, text="➕ 選取項目名額+1", command=lambda: self.adjust_selected_quota(1))
        btn_plus.pack(side=tk.LEFT, padx=2)

        btn_minus = ttk.Button(act_frame, text="➖ 選取項目名額-1", command=lambda: self.adjust_selected_quota(-1))
        btn_minus.pack(side=tk.LEFT, padx=2)

        btn_del = ttk.Button(act_frame, text="🗑️ 移除選取項目", command=self.delete_selected_item)
        btn_del.pack(side=tk.RIGHT, padx=2)

        # === 右側：名單輸入與統計 ===
        right_frame = ttk.LabelFrame(main_paned, text="👥 抽籤名單輸入", padding=10)
        main_paned.add(right_frame, weight=2)

        quick_num_frame = ttk.Frame(right_frame)
        quick_num_frame.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(quick_num_frame, text="⚡ 快速填入座號:").pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(quick_num_frame, text="1~15號", width=6, command=lambda: self.quick_numbers(15)).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick_num_frame, text="1~30號", width=6, command=lambda: self.quick_numbers(30)).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick_num_frame, text="清空", width=5, command=self.clear_names).pack(side=tk.RIGHT, padx=2)

        self.txt_names = tk.Text(right_frame, width=28, height=14, font=("Microsoft JhengHei", 10))
        self.txt_names.pack(fill=tk.BOTH, expand=True)
        self.txt_names.bind("<KeyRelease>", lambda e: self.update_stats())

        # 名單統計狀態
        self.lbl_names_stat = ttk.Label(right_frame, text="已輸入: 0 人", font=("Microsoft JhengHei", 9, "bold"))
        self.lbl_names_stat.pack(anchor=tk.W, pady=(4, 2))

        self.lbl_status = ttk.Label(right_frame, text="", wraplength=260, font=("Microsoft JhengHei", 9))
        self.lbl_status.pack(anchor=tk.W, pady=(0, 4))

        # 抽籤大按鈕
        btn_lottery = tk.Button(right_frame, text="🎉 開始隨機抽籤", font=("Microsoft JhengHei", 12, "bold"),
                                bg="#2563eb", fg="white", activebackground="#1d4ed8", activeforeground="white",
                                relief=tk.RAISED, cursor="hand2", command=self.draw_lottery, pady=8)
        btn_lottery.pack(fill=tk.X, pady=(6, 0))

        # === 底部：抽籤結果顯示 ===
        bot_frame = ttk.LabelFrame(self, text="🏆 抽籤分配結果", padding=10)
        bot_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=(5, 10))

        res_tree_frame = ttk.Frame(bot_frame)
        res_tree_frame.pack(fill=tk.BOTH, expand=True)

        self.res_tree = ttk.Treeview(res_tree_frame, columns=("category", "name", "quota", "assigned"), show="headings")
        self.res_tree.heading("category", text="分類")
        self.res_tree.heading("name", text="工作項目")
        self.res_tree.heading("quota", text="名額")
        self.res_tree.heading("assigned", text="抽中人員名單")

        self.res_tree.column("category", width=90, anchor=tk.CENTER)
        self.res_tree.column("name", width=140, anchor=tk.W)
        self.res_tree.column("quota", width=60, anchor=tk.CENTER)
        self.res_tree.column("assigned", width=450, anchor=tk.W)

        res_scroll = ttk.Scrollbar(res_tree_frame, orient=tk.VERTICAL, command=self.res_tree.yview)
        self.res_tree.configure(yscrollcommand=res_scroll.set)

        self.res_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        res_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 結果列操作 (複製、匯出)
        res_act_frame = ttk.Frame(bot_frame)
        res_act_frame.pack(fill=tk.X, pady=(6, 0))

        self.lbl_unassigned = ttk.Label(res_act_frame, text="", foreground="#d97706", font=("Microsoft JhengHei", 9, "bold"))
        self.lbl_unassigned.pack(side=tk.LEFT, padx=4)

        ttk.Button(res_act_frame, text="📊 匯出 CSV", command=self.export_csv).pack(side=tk.RIGHT, padx=4)
        ttk.Button(res_act_frame, text="📋 複製文字結果", command=self.copy_results).pack(side=tk.RIGHT, padx=4)

    def load_preset(self, preset_key):
        tpl = PRESETS.get(preset_key)
        if not tpl:
            return

        for item in self.tree.get_children():
            self.tree.delete(item)

        for cat, name, desc, quota in tpl["items"]:
            self.tree.insert("", tk.END, values=(cat, name, desc, quota))

        self.update_stats()

    def add_item(self):
        cat = self.entry_cat.get().strip()
        name = self.entry_name.get().strip()
        desc = self.entry_desc.get().strip()
        try:
            quota = int(self.spin_quota.get())
            if quota < 1:
                quota = 1
        except ValueError:
            quota = 1

        if not name:
            messagebox.showwarning("提示", "請輸入項目名稱！")
            self.entry_name.focus()
            return

        self.tree.insert("", tk.END, values=(cat, name, desc, quota))
        self.entry_name.delete(0, tk.END)
        self.entry_desc.delete(0, tk.END)
        self.spin_quota.set(1)
        self.update_stats()

    def delete_selected_item(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("提示", "請先點選欲刪除的項目！")
            return
        for s in sel:
            self.tree.delete(s)
        self.update_stats()

    def adjust_selected_quota(self, delta):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("提示", "請先在列表中點選一個項目！")
            return
        for s in sel:
            cat, name, desc, q = self.tree.item(s, "values")
            new_q = max(1, int(q) + delta)
            self.tree.item(s, values=(cat, name, desc, new_q))
        self.update_stats()

    def get_candidate_names(self):
        content = self.txt_names.get("1.0", tk.END).strip()
        if not content:
            return []
        import re
        tokens = re.split(r'[\r\n,;，；\t]+', content)
        res = []
        for t in tokens:
            t = t.strip()
            if t:
                # also handle spaces
                parts = t.split()
                for p in parts:
                    if p:
                        res.push(p) if hasattr(res, 'push') else res.append(p)
        return res

    def quick_numbers(self, n):
        nums = [f"{i}號" for i in range(1, n + 1)]
        self.txt_names.delete("1.0", tk.END)
        self.txt_names.insert("1.0", "\n".join(nums))
        self.update_stats()

    def clear_names(self):
        self.txt_names.delete("1.0", tk.END)
        self.update_stats()

    def get_total_quota(self):
        total = 0
        for item in self.tree.get_children():
            _, _, _, q = self.tree.item(item, "values")
            total += int(q)
        return total

    def update_stats(self):
        candidates = self.get_candidate_names()
        c_count = len(candidates)
        items_count = len(self.tree.get_children())
        total_quota = self.get_total_quota()

        self.quota_summary_var.set(f"項目總計：{items_count} 項 | 需求總人數：{total_quota} 人")
        self.lbl_names_stat.config(text=f"已輸入名單：{c_count} 人")

        diff = c_count - total_quota
        if c_count == 0 and total_quota == 0:
            self.lbl_status.config(text="", foreground="#555")
        elif diff == 0:
            self.lbl_status.config(text="✅ 人數與名額剛好吻合！", foreground="#16a34a")
        elif diff > 0:
            self.lbl_status.config(text=f"⚠️ 名單多 {diff} 人，將列為未分配/候補。", foreground="#d97706")
        else:
            self.lbl_status.config(text=f"ℹ️ 尚缺 {abs(diff)} 人，部分名額將留空。", foreground="#dc2626")

    def draw_lottery(self):
        items = self.tree.get_children()
        if not items:
            messagebox.showwarning("提示", "目前沒有任何工作項目，請先新增或載入範本！")
            return

        candidates = self.get_candidate_names()
        if not candidates:
            messagebox.showwarning("提示", "請輸入抽籤名單或點選快速填入座號！")
            self.txt_names.focus()
            return

        # 隨機洗牌
        shuffled = list(candidates)
        random.shuffle(shuffled)

        for r in self.res_tree.get_children():
            self.res_tree.delete(r)

        ptr = 0
        self.last_results = []

        for item_id in items:
            cat, name, desc, q_str = self.tree.item(item_id, "values")
            quota = int(q_str)
            assigned = []
            for _ in range(quota):
                if ptr < len(shuffled):
                    assigned.append(shuffled[ptr])
                    ptr += 1
                else:
                    assigned.append("（缺額）")

            assigned_str = "、".join(assigned)
            item_display_name = name + (f" ({desc})" if desc else "")
            self.res_tree.insert("", tk.END, values=(cat, item_display_name, quota, assigned_str))
            self.last_results.append({
                "category": cat,
                "name": name,
                "desc": desc,
                "quota": quota,
                "assigned": assigned
            })

        self.last_unassigned = shuffled[ptr:]
        if self.last_unassigned:
            self.lbl_unassigned.config(text=f"🌟 未分配/候補人員 ({len(self.last_unassigned)}人): " + "、".join(self.last_unassigned))
        else:
            self.lbl_unassigned.config(text="✨ 所有名單人員皆已成功分配！")

        messagebox.showinfo("抽籤完成", "🎉 隨機抽籤已完成！結果已更新於下方表格。")

    def copy_results(self):
        if not self.last_results:
            messagebox.showinfo("提示", "請先進行抽籤再複製結果！")
            return

        lines = [f"【{self.preset_var.get()} - 抽籤分配結果】", f"抽籤時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", "-" * 40]
        for r in self.last_results:
            cat_str = f"[{r['category']}] " if r['category'] else ""
            desc_str = f" ({r['desc']})" if r['desc'] else ""
            names_str = "、".join(r['assigned'])
            lines.append(f"{cat_str}{r['name']}{desc_str} ({r['quota']}人)：{names_str}")

        if self.last_unassigned:
            lines.append("-" * 40)
            lines.append(f"【未分配/候補人員 ({len(self.last_unassigned)}人)】：" + "、".join(self.last_unassigned))

        full_text = "\n".join(lines)
        self.clipboard_clear()
        self.clipboard_append(full_text)
        messagebox.showinfo("複製成功", "✅ 抽籤結果已複製到剪貼簿，可直接貼上至 LINE 或 Word！")

    def export_csv(self):
        if not self.last_results:
            messagebox.showinfo("提示", "請先進行抽籤再匯出！")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV 檔案", "*.csv"), ("所有檔案", "*.*")],
            initialfile=f"抽籤分配表_{datetime.now().strftime('%Y%m%d')}.csv"
        )
        if not filepath:
            return

        try:
            with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(["分組", "項目名稱", "備註說明", "需求人數", "抽中分配名單"])
                for r in self.last_results:
                    writer.writerow([r["category"], r["name"], r["desc"], r["quota"], "、".join(r["assigned"])])
                if self.last_unassigned:
                    writer.writerow(["未分配", "候補/輪空", "", len(self.last_unassigned), "、".join(self.last_unassigned)])
            messagebox.showinfo("匯出成功", f"檔案已成功儲存至：\n{filepath}")
        except Exception as e:
            messagebox.showerror("錯誤", f"儲存失敗：{e}")

if __name__ == "__main__":
    app = LotteryApp()
    app.mainloop()
