import os
import sys
import threading
import subprocess
import shutil
import time
import tkinter as tk
from tkinter import filedialog, messagebox

# Optional: for cleaning MIDI
try:
    import pretty_midi
    import numpy as np
    PRETTY_MIDI_OK = True
except Exception:
    PRETTY_MIDI_OK = False


def which(cmd: str) -> str | None:
    return shutil.which(cmd)


def run_cmd(cmd_list, log_fn, cwd=None):
    log_fn(f"$ {' '.join(cmd_list)}")
    p = subprocess.Popen(
        cmd_list,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        cwd=cwd
    )
    for line in p.stdout:
        log_fn(line.rstrip("\n"))
    p.wait()
    return p.returncode


def quantize(x, grid):
    return float(np.round(x / grid) * grid)


def extract_top_melody(pm: "pretty_midi.PrettyMIDI"):
    notes = []
    for inst in pm.instruments:
        if inst.is_drum:
            continue
        notes.extend(inst.notes)

    if not notes:
        return []

    notes.sort(key=lambda n: (n.start, -n.pitch))

    melody = []
    i = 0
    while i < len(notes):
        start = notes[i].start
        group = [notes[i]]
        i += 1
        while i < len(notes) and abs(notes[i].start - start) <= 0.02:
            group.append(notes[i])
            i += 1
        top = max(group, key=lambda n: n.pitch)
        melody.append(top)
    return melody


def clean_midi_for_game(input_mid, output_mid, bpm=120, grid_div=2, min_note_sec=0.08):
    """
    grid_div=2 -> 1/8 拍
    grid_div=4 -> 1/16 拍
    """
    if not PRETTY_MIDI_OK:
        raise RuntimeError("pretty_midi / numpy 未安装，无法进行清理。")

    pm = pretty_midi.PrettyMIDI(input_mid)
    grid = (60.0 / float(bpm)) / float(grid_div)

    melody_notes = extract_top_melody(pm)

    out = pretty_midi.PrettyMIDI(initial_tempo=float(bpm))
    inst = pretty_midi.Instrument(program=0, name="Melody")

    for n in melody_notes:
        s = quantize(n.start, grid)
        e = quantize(n.end, grid)
        if e <= s:
            e = s + grid
        if (e - s) < float(min_note_sec):
            continue

        inst.notes.append(pretty_midi.Note(
            velocity=int(np.clip(n.velocity, 20, 110)),
            pitch=int(n.pitch),
            start=float(s),
            end=float(e),
        ))

    out.instruments.append(inst)
    out.write(output_mid)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MP3 -> MIDI (燕云十六声用)")

        self.in_path = tk.StringVar(value="")
        self.out_dir = tk.StringVar(value=os.path.abspath("./output"))
        self.mode = tk.StringVar(value="vocals")  # vocals or instrumental
        self.do_clean = tk.BooleanVar(value=True)
        self.bpm = tk.StringVar(value="120")
        self.grid_div = tk.StringVar(value="2")  # 2=1/8, 4=1/16
        self.status = tk.StringVar(value="就绪")

        self._build_ui()

        os.makedirs(self.out_dir.get(), exist_ok=True)

    def _build_ui(self):
        pad = 8

        frm1 = tk.Frame(self)
        frm1.pack(fill="x", padx=pad, pady=(pad, 0))

        tk.Label(frm1, text="输入文件 (MP3/WAV):").grid(row=0, column=0, sticky="w")
        tk.Entry(frm1, textvariable=self.in_path, width=60).grid(row=0, column=1, padx=(6, 6))
        tk.Button(frm1, text="选择文件", command=self.pick_file).grid(row=0, column=2)

        frm2 = tk.Frame(self)
        frm2.pack(fill="x", padx=pad, pady=(pad, 0))

        tk.Label(frm2, text="输出文件夹:").grid(row=0, column=0, sticky="w")
        tk.Entry(frm2, textvariable=self.out_dir, width=60).grid(row=0, column=1, padx=(6, 6))
        tk.Button(frm2, text="选择文件夹", command=self.pick_outdir).grid(row=0, column=2)

        frm3 = tk.LabelFrame(self, text="转换模式")
        frm3.pack(fill="x", padx=pad, pady=(pad, 0))

        tk.Radiobutton(frm3, text="人声转换 (demucs 分离后转 MIDI)", variable=self.mode, value="vocals").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        tk.Radiobutton(frm3, text="纯音乐转换 (直接转 MIDI)", variable=self.mode, value="instrumental").grid(row=1, column=0, sticky="w", padx=6, pady=4)

        frm4 = tk.LabelFrame(self, text="燕云优化 (推荐开启)")
        frm4.pack(fill="x", padx=pad, pady=(pad, 0))

        tk.Checkbutton(frm4, text="清理成单旋律 + 量化节奏", variable=self.do_clean).grid(row=0, column=0, sticky="w", padx=6, pady=4)

        sub = tk.Frame(frm4)
        sub.grid(row=1, column=0, sticky="w", padx=6, pady=(0, 6))

        tk.Label(sub, text="BPM:").grid(row=0, column=0, sticky="w")
        tk.Entry(sub, textvariable=self.bpm, width=6).grid(row=0, column=1, padx=(6, 16))

        tk.Label(sub, text="量化网格:").grid(row=0, column=2, sticky="w")
        grid_menu = tk.OptionMenu(sub, self.grid_div, "2", "4")
        grid_menu.config(width=6)
        grid_menu.grid(row=0, column=3, padx=(6, 0))
        tk.Label(sub, text="2=1/8拍, 4=1/16拍").grid(row=0, column=4, padx=(10, 0))

        frm5 = tk.Frame(self)
        frm5.pack(fill="x", padx=pad, pady=(pad, 0))

        tk.Button(frm5, text="开始转换", command=self.start).pack(side="left")
        tk.Label(frm5, textvariable=self.status).pack(side="left", padx=(10, 0))

        frm6 = tk.Frame(self)
        frm6.pack(fill="both", expand=True, padx=pad, pady=pad)

        tk.Label(frm6, text="日志:").pack(anchor="w")
        self.log = tk.Text(frm6, height=18, wrap="word")
        self.log.pack(fill="both", expand=True)

        hint = tk.Label(self, text="提示: 也可以把文件路径复制到输入框里。", fg="gray")
        hint.pack(anchor="w", padx=pad, pady=(0, pad))

    def log_line(self, s: str):
        self.log.insert("end", s + "\n")
        self.log.see("end")
        self.update_idletasks()

    def pick_file(self):
        p = filedialog.askopenfilename(
            title="选择音频文件",
            filetypes=[("Audio", "*.mp3 *.wav *.m4a *.flac"), ("All", "*.*")]
        )
        if p:
            self.in_path.set(p)

    def pick_outdir(self):
        d = filedialog.askdirectory(title="选择输出文件夹")
        if d:
            self.out_dir.set(d)
            os.makedirs(d, exist_ok=True)

    def start(self):
        in_file = self.in_path.get().strip()
        out_dir = self.out_dir.get().strip()

        if not in_file or not os.path.exists(in_file):
            messagebox.showerror("错误", "请选择一个存在的 MP3/WAV 文件。")
            return

        if not out_dir:
            messagebox.showerror("错误", "请选择输出文件夹。")
            return

        os.makedirs(out_dir, exist_ok=True)

        # Basic dependency checks
        if which("basic-pitch") is None:
            messagebox.showerror("缺少依赖", "找不到 basic-pitch 命令。请先执行: pip install -U basic-pitch")
            return

        if self.mode.get() == "vocals" and which("demucs") is None:
            messagebox.showerror("缺少依赖", "找不到 demucs 命令。请先执行: pip install -U demucs")
            return

        if self.do_clean.get() and not PRETTY_MIDI_OK:
            messagebox.showwarning("提示", "未检测到 pretty_midi/numpy，无法进行清理。你仍可先输出原始 MIDI。\n安装: pip install -U pretty_midi numpy")

        self.status.set("处理中...")
        self.log.delete("1.0", "end")

        t = threading.Thread(target=self._worker, daemon=True)
        t.start()

    def _worker(self):
        try:
            in_file = self.in_path.get().strip()
            out_dir = self.out_dir.get().strip()
            base = os.path.splitext(os.path.basename(in_file))[0]

            work_dir = os.path.join(out_dir, f"_work_{base}_{int(time.time())}")
            os.makedirs(work_dir, exist_ok=True)

            self.log_line(f"输入: {in_file}")
            self.log_line(f"输出: {out_dir}")
            self.log_line(f"工作目录: {work_dir}")
            self.log_line("")

            target_audio = in_file

            # Mode: vocals
            if self.mode.get() == "vocals":
                self.log_line("模式: 人声转换, 正在用 demucs 分离人声...")
                demucs_out = os.path.join(work_dir, "separated")
                os.makedirs(demucs_out, exist_ok=True)

                code = run_cmd(
                    ["demucs", "-n", "htdemucs", "--two-stems=vocals", "-o", demucs_out, in_file],
                    self.log_line
                )
                if code != 0:
                    raise RuntimeError("demucs 运行失败，请看日志。")

                # find vocals.wav
                vocals_wav = None
                # typical path: separated/htdemucs/<base>/vocals.wav
                candidate = os.path.join(demucs_out, "htdemucs", base, "vocals.wav")
                if os.path.exists(candidate):
                    vocals_wav = candidate
                else:
                    # fallback scan
                    for root, _, files in os.walk(demucs_out):
                        if "vocals.wav" in files:
                            vocals_wav = os.path.join(root, "vocals.wav")
                            break

                if not vocals_wav:
                    raise RuntimeError("找不到 demucs 输出的 vocals.wav，请看日志。")

                target_audio = vocals_wav
                self.log_line(f"人声文件: {target_audio}")
                self.log_line("")

            # Run basic-pitch
            self.log_line("正在用 basic-pitch 转 MIDI...")
            bp_out = os.path.join(work_dir, "basic_pitch_out")
            os.makedirs(bp_out, exist_ok=True)

            code = run_cmd(
                ["basic-pitch", bp_out, target_audio],
                self.log_line
            )
            if code != 0:
                raise RuntimeError("basic-pitch 运行失败，请看日志。")

            # find midi output
            raw_mid = None
            # basic-pitch often outputs <audio_basename>.mid
            for f in os.listdir(bp_out):
                if f.lower().endswith(".mid") or f.lower().endswith(".midi"):
                    raw_mid = os.path.join(bp_out, f)
                    break

            if not raw_mid:
                # scan deeper just in case
                for root, _, files in os.walk(bp_out):
                    for f in files:
                        if f.lower().endswith(".mid") or f.lower().endswith(".midi"):
                            raw_mid = os.path.join(root, f)
                            break
                    if raw_mid:
                        break

            if not raw_mid:
                raise RuntimeError("找不到 basic-pitch 输出的 MIDI 文件，请看日志。")

            final_raw = os.path.join(out_dir, f"{base}_raw.mid")
            shutil.copy2(raw_mid, final_raw)
            self.log_line("")
            self.log_line(f"已输出原始 MIDI: {final_raw}")

            # Cleaning for game
            if self.do_clean.get() and PRETTY_MIDI_OK:
                bpm = int(self.bpm.get().strip() or "120")
                grid_div = int(self.grid_div.get().strip() or "2")
                cleaned = os.path.join(out_dir, f"{base}_yanyun_clean.mid")
                self.log_line("")
                self.log_line("正在进行燕云清理: 单旋律 + 节奏量化...")
                clean_midi_for_game(final_raw, cleaned, bpm=bpm, grid_div=grid_div)
                self.log_line(f"已输出清理后 MIDI: {cleaned}")

            self.status.set("完成 ✅")
            self.log_line("")
            self.log_line("完成。")

        except Exception as e:
            self.status.set("失败 ❌")
            self.log_line("")
            self.log_line(f"错误: {e}")
            messagebox.showerror("转换失败", str(e))


if __name__ == "__main__":
    app = App()
    app.geometry("860x520")
    app.mainloop()
