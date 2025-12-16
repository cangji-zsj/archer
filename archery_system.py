import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import datetime
import os


class ArcheryBracketSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("射箭比赛编排系统 v1.0")
        self.root.geometry("900x700")
        self.root.resizable(True, True)

        self.data = None
        self.brackets = []

        self.setup_ui()

    def setup_ui(self):
        # 标题
        title_frame = ttk.Frame(self.root, padding="10")
        title_frame.pack(fill=tk.X)

        title_label = ttk.Label(
            title_frame,
            text="🎯 射箭比赛编排系统",
            font=("Arial", 20, "bold")
        )
        title_label.pack()

        subtitle = ttk.Label(
            title_frame,
            text="支持个人赛/团体赛 | 自动生成对阵表与靶位分配",
            font=("Arial", 10)
        )
        subtitle.pack()

        # 上传区域
        upload_frame = ttk.LabelFrame(self.root, text="📁 上传排位赛成绩", padding="15")
        upload_frame.pack(fill=tk.X, padx=20, pady=10)

        instruction = ttk.Label(
            upload_frame,
            text="请上传Excel表格，格式：第一列=排名，第二列=姓名",
            foreground="gray"
        )
        instruction.pack(anchor=tk.W)

        btn_frame = ttk.Frame(upload_frame)
        btn_frame.pack(fill=tk.X, pady=5)

        self.upload_btn = ttk.Button(
            btn_frame,
            text="选择文件",
            command=self.load_file,
            width=15
        )
        self.upload_btn.pack(side=tk.LEFT, padx=5)

        self.file_label = ttk.Label(btn_frame, text="未选择文件", foreground="gray")
        self.file_label.pack(side=tk.LEFT, padx=10)

        # 比赛类型选择
        type_frame = ttk.LabelFrame(self.root, text="⚙️ 比赛设置", padding="15")
        type_frame.pack(fill=tk.X, padx=20, pady=10)

        ttk.Label(type_frame, text="比赛类型：").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.match_type = tk.StringVar(value="individual")
        ttk.Radiobutton(
            type_frame,
            text="个人赛 (每局3箭)",
            variable=self.match_type,
            value="individual"
        ).grid(row=0, column=1, sticky=tk.W, padx=10)
        ttk.Radiobutton(
            type_frame,
            text="团体赛 (每局6箭)",
            variable=self.match_type,
            value="team"
        ).grid(row=0, column=2, sticky=tk.W, padx=10)

        # 生成按钮
        self.generate_btn = ttk.Button(
            type_frame,
            text="🎯 生成对阵编排",
            command=self.generate_brackets,
            state=tk.DISABLED
        )
        self.generate_btn.grid(row=1, column=0, columnspan=3, pady=10)

        # 结果显示区域
        result_frame = ttk.LabelFrame(self.root, text="📋 对阵编排结果", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # 添加滚动条
        tree_scroll = ttk.Scrollbar(result_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree = ttk.Treeview(
            result_frame,
            columns=("round", "match", "left", "vs", "right", "target", "color"),
            show="headings",
            yscrollcommand=tree_scroll.set,
            height=15
        )
        tree_scroll.config(command=self.tree.yview)

        # 定义列
        self.tree.heading("round", text="轮次")
        self.tree.heading("match", text="场次")
        self.tree.heading("left", text="左侧选手(A靶)")
        self.tree.heading("vs", text="")
        self.tree.heading("right", text="右侧选手(B靶)")
        self.tree.heading("target", text="靶位")
        self.tree.heading("color", text="颜色标识")

        self.tree.column("round", width=100, anchor=tk.CENTER)
        self.tree.column("match", width=80, anchor=tk.CENTER)
        self.tree.column("left", width=150, anchor=tk.CENTER)
        self.tree.column("vs", width=40, anchor=tk.CENTER)
        self.tree.column("right", width=150, anchor=tk.CENTER)
        self.tree.column("target", width=100, anchor=tk.CENTER)
        self.tree.column("color", width=120, anchor=tk.CENTER)

        self.tree.pack(fill=tk.BOTH, expand=True)

        # 导出按钮
        export_frame = ttk.Frame(self.root)
        export_frame.pack(fill=tk.X, padx=20, pady=10)

        self.export_btn = ttk.Button(
            export_frame,
            text="💾 导出为Excel",
            command=self.export_results,
            state=tk.DISABLED
        )
        self.export_btn.pack(side=tk.LEFT, padx=5)

        self.print_btn = ttk.Button(
            export_frame,
            text="🖨️ 导出为TXT",
            command=self.export_txt,
            state=tk.DISABLED
        )
        self.print_btn.pack(side=tk.LEFT, padx=5)

        # 版权信息
        footer = ttk.Label(
            self.root,
            text="© 2024 射箭比赛编排系统 | 严格遵循国际射箭联合会规则",
            font=("Arial", 8),
            foreground="gray"
        )
        footer.pack(pady=5)

    def load_file(self):
        filepath = filedialog.askopenfilename(
            title="选择排位赛成绩表",
            filetypes=[
                ("Excel文件", "*.xlsx *.xls"),
                ("CSV文件", "*.csv"),
                ("所有文件", "*.*")
            ]
        )

        if not filepath:
            return

        try:
            if filepath.endswith('.csv'):
                self.data = pd.read_csv(filepath, header=None)
            else:
                self.data = pd.read_excel(filepath, header=None)

            # 验证数据格式
            if len(self.data.columns) < 2:
                raise ValueError("表格至少需要两列：排名和姓名")

            self.data.columns = ['排名', '姓名'] + [f'列{i}' for i in range(2, len(self.data.columns))]
            self.data = self.data[['排名', '姓名']]

            # 显示文件名
            filename = os.path.basename(filepath)
            self.file_label.config(
                text=f"✓ {filename} ({len(self.data)}名选手)",
                foreground="green"
            )

            self.generate_btn.config(state=tk.NORMAL)

        except Exception as e:
            messagebox.showerror("错误", f"文件读取失败：{str(e)}")
            self.file_label.config(text="文件格式错误", foreground="red")

    def generate_brackets(self):
        if self.data is None:
            messagebox.showwarning("提示", "请先上传排位赛成绩表")
            return

        num_players = len(self.data)

        # 检查是否为2的幂次
        if num_players & (num_players - 1) != 0:
            messagebox.showwarning(
                "提示",
                f"当前选手数量为{num_players}人，不是2的幂次。\n"
                f"建议选手数量为：8, 16, 32, 64等。\n"
                f"系统将为前{self.get_valid_bracket_size(num_players)}名选手生成对阵表。"
            )
            num_players = self.get_valid_bracket_size(num_players)
            self.data = self.data.head(num_players)

        self.brackets = []
        self.tree.delete(*self.tree.get_children())

        # 生成首轮对阵
        first_round = []
        for i in range(num_players // 2):
            left_rank = i + 1
            right_rank = num_players - i

            left_player = self.data[self.data['排名'] == left_rank]['姓名'].values[0]
            right_player = self.data[self.data['排名'] == right_rank]['姓名'].values[0]

            match = {
                'round': f'1/{num_players // 2}决赛',
                'match_num': i + 1,
                'left': f"#{left_rank} {left_player}",
                'right': f"#{right_rank} {right_player}",
                'target': f"{i + 1}号靶",
                'left_color': '🟢 绿色',
                'right_color': '🔴 红色'
            }
            first_round.append(match)

        self.brackets.extend(first_round)

        # 显示在表格中
        for match in self.brackets:
            self.tree.insert('', tk.END, values=(
                match['round'],
                f"第{match['match_num']}场",
                match['left'],
                "VS",
                match['right'],
                match['target'],
                f"{match['left_color']} vs {match['right_color']}"
            ))

        self.export_btn.config(state=tk.NORMAL)
        self.print_btn.config(state=tk.NORMAL)

        messagebox.showinfo(
            "成功",
            f"对阵编排完成！\n\n"
            f"• 参赛人数：{num_players}人\n"
            f"• 首轮场次：{len(first_round)}场\n"
            f"• 比赛类型：{'个人赛' if self.match_type.get() == 'individual' else '团体赛'}\n\n"
            f"上半区种子：#{1}\n"
            f"下半区种子：#{2}"
        )

    def get_valid_bracket_size(self, n):
        powers = [8, 16, 32, 64, 128]
        for p in powers:
            if n <= p:
                return p
        return 128

    def export_results(self):
        if not self.brackets:
            messagebox.showwarning("提示", "请先生成对阵编排")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")],
            initialfile=f"射箭对阵表_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        if not filepath:
            return

        try:
            export_data = []
            for match in self.brackets:
                export_data.append({
                    '轮次': match['round'],
                    '场次': f"第{match['match_num']}场",
                    '左侧选手(A靶)': match['left'],
                    '右侧选手(B靶)': match['right'],
                    '靶位': match['target'],
                    '左侧颜色': match['left_color'],
                    '右侧颜色': match['right_color']
                })

            df = pd.DataFrame(export_data)
            df.to_excel(filepath, index=False)

            messagebox.showinfo("成功", f"对阵表已导出至：\n{filepath}")

        except Exception as e:
            messagebox.showerror("错误", f"导出失败：{str(e)}")

    def export_txt(self):
        if not self.brackets:
            messagebox.showwarning("提示", "请先生成对阵编排")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt")],
            initialfile=f"射箭对阵表_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        )

        if not filepath:
            return

        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write("=" * 80 + "\n")
                f.write("射箭比赛对阵编排表\n".center(76))
                f.write(f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n".center(76))
                f.write("=" * 80 + "\n\n")

                for match in self.brackets:
                    f.write(f"【{match['round']}】第{match['match_num']}场\n")
                    f.write(f"  {match['target']}\n")
                    f.write(f"  A靶 {match['left_color']}：{match['left']}\n")
                    f.write(f"       VS\n")
                    f.write(f"  B靶 {match['right_color']}：{match['right']}\n")
                    f.write("-" * 80 + "\n\n")

            messagebox.showinfo("成功", f"对阵表已导出至：\n{filepath}")

        except Exception as e:
            messagebox.showerror("错误", f"导出失败：{str(e)}")


def main():
    root = tk.Tk()
    app = ArcheryBracketSystem(root)
    root.mainloop()


if __name__ == "__main__":
    main()