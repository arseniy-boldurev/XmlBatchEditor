#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
XML Batch Editor — MVP

Что умеет:
- Открыть архив (.zip) с XML файлами (вложенные папки поддерживаются).
- Применить набор правил замены:
    • По имени тега (без учёта пространств имён)
    • По XPath-выражению (расширенно)
- Записать новый архив со всеми файлами, где XML — изменены, остальные — скопированы как есть.
- Работает БЕЗ распаковки на диск → нет проблем с «слишком длинным путём» в Windows.
- Лог выполнения + CSV-отчёт об изменениях (в архиве рядом с результатом).
- Сохранение/загрузка пресетов правил (JSON).

Собрать .exe: 
    pyinstaller --noconfirm --onefile --windowed --name XMLBatchEditor app.py

После сборки можно создать ярлык на рабочем столе и перетаскивать на него .zip архивы —
приложение прочитает путь из аргументов командной строки и подставит в поле «Входной архив».

Зависимости:
    pip install lxml ttkbootstrap

Автор: Работаем братья ✦
"""

import io
import json
import sys
import csv
import time
import zipfile
from dataclasses import dataclass, asdict
from pathlib import PurePosixPath
from typing import List, Optional

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

try:
    import ttkbootstrap as tb
    from ttkbootstrap.constants import *  # noqa
    TKROOT = tb.Window(themename="cosmo")
except Exception:
    # На случай, если ttkbootstrap не установлен — использовать стандартный Tk
    TKROOT = tk.Tk()

from lxml import etree

APP_TITLE = "XML Batch Editor — MVP"
VERSION = "0.3.0"

# ---------------------------- Модель правил ---------------------------- #
@dataclass
class Rule:
    mode: str        # 'tag' | 'xpath'
    pattern: str     # имя тега или XPath
    new_value: str   # чем заменить .text

    def to_row(self):
        return (self.mode, self.pattern, self.new_value)

    @staticmethod
    def from_row(row):
        return Rule(mode=row[0], pattern=row[1], new_value=row[2])

# ------------------------- XML утилиты/ядро --------------------------- #

NS_AWARE_HELP = (
    "Подсказка по тегам и NS: режим ‘По имени тега’ ищет по local-name(),\n"
    "то есть <ns:C_STI_ORIG> тоже будет найден. Для сложных случаев используйте XPath."
)

PARSER = etree.XMLParser(remove_blank_text=False, recover=True, strip_cdata=False)


def _elements_by_tag_localname(root: etree._Element, localname: str):
    # XPath по local-name() — чтобы игнорировать пространства имён
    expr = f".//*[local-name()='{localname}']"
    return root.xpath(expr)


def apply_rules(xml_bytes: bytes, rules: List[Rule]):
    """Возвращает (new_bytes, applied_count, errors).
    Не меняет структуру, сохраняет исходную кодировку из декларации, если возможно.
    """
    errors: List[str] = []
    applied = 0

    try:
        root = etree.fromstring(xml_bytes, parser=PARSER)
    except Exception as e:
        errors.append(f"parse error: {e}")
        return xml_bytes, applied, errors

    for rule in rules:
        try:
            if rule.mode == 'tag':
                els = _elements_by_tag_localname(root, rule.pattern)
            else:  # xpath
                els = root.xpath(rule.pattern)
                # отфильтровать только элементы
                els = [el for el in els if isinstance(el, etree._Element)]

            for el in els:
                el.text = rule.new_value
                applied += 1
        except Exception as e:
            errors.append(f"rule '{rule.mode}:{rule.pattern}': {e}")

    # Пытаемся сохранить в исходной кодировке (при её наличии)
    decl_encoding = _extract_declared_encoding(xml_bytes)
    encoding = decl_encoding or 'utf-8'

    try:
        new_bytes = etree.tostring(
            root,
            encoding=encoding,
            xml_declaration=True,
            pretty_print=False,
            with_tail=True,
        )
    except Exception as e:
        errors.append(f"serialize error: {e}")
        # в крайнем случае — отдать исходник
        return xml_bytes, applied, errors

    return new_bytes, applied, errors


def _extract_declared_encoding(data: bytes) -> Optional[str]:
    head = data[:128].decode('ascii', errors='ignore')
    if "<?xml" in head and "encoding" in head:
        # примитивный парсер декларации
        try:
            start = head.index("encoding")
            q1 = head.index('"', start)
            q2 = head.index('"', q1 + 1)
            return head[q1 + 1:q2]
        except Exception:
            pass
    return None

# --------------------------- ZIP обработка ---------------------------- #

@dataclass
class ProcessStats:
    total_files: int = 0
    xml_changed: int = 0
    xml_unchanged: int = 0
    copied_other: int = 0
    errors: int = 0


def process_zip(in_zip_path: str, out_zip_path: str, rules: List[Rule], log_writer=None) -> ProcessStats:
    stats = ProcessStats()
    start = time.time()

    with zipfile.ZipFile(in_zip_path, 'r') as zin, zipfile.ZipFile(out_zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            stats.total_files += 1
            arcname = info.filename  # тип: str с POSIX-разделителями

            try:
                data = zin.read(info)
                if arcname.lower().endswith('.xml'):
                    new_bytes, applied, errors = apply_rules(data, rules)
                    if errors and log_writer:
                        for err in errors:
                            log_writer.writerow([arcname, 'error', err])
                        stats.errors += len(errors)

                    if applied > 0 and new_bytes != data:
                        stats.xml_changed += 1
                        _write_clone_info(zout, info, new_bytes)
                        if log_writer:
                            log_writer.writerow([arcname, 'changed', f"applied={applied}"])
                    else:
                        stats.xml_unchanged += 1
                        _write_clone_info(zout, info, data)
                        if log_writer:
                            log_writer.writerow([arcname, 'unchanged', f"applied={applied}"])
                else:
                    stats.copied_other += 1
                    _write_clone_info(zout, info, data)
            except Exception as e:
                stats.errors += 1
                if log_writer:
                    log_writer.writerow([arcname, 'fatal', str(e)])

    dur = time.time() - start
    if log_writer:
        log_writer.writerow(["__SUMMARY__", json.dumps(asdict(stats), ensure_ascii=False), f"{dur:.2f}s"])
    return stats


def _write_clone_info(zout: zipfile.ZipFile, src_info: zipfile.ZipInfo, data: bytes):
    # Сохранить метаданные по максимуму
    zi = zipfile.ZipInfo(filename=src_info.filename, date_time=src_info.date_time)
    zi.compress_type = zipfile.ZIP_DEFLATED
    zi.create_system = src_info.create_system
    zi.external_attr = src_info.external_attr
    zout.writestr(zi, data)

# ------------------------------- GUI --------------------------------- #

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"{APP_TITLE} — v{VERSION}")
        self.root.geometry("900x620")

        self.in_path = tk.StringVar()
        self.out_path = tk.StringVar()

        self._build_ui()
        self._maybe_load_cli_arg()

    # --- UI construction ---
    def _build_ui(self):
        pad = 10

        frm_top = ttk.Frame(self.root, padding=pad)
        frm_top.pack(fill='x')

        # Input ZIP
        ttk.Label(frm_top, text="Входной архив (.zip):").grid(row=0, column=0, sticky='w')
        ttk.Entry(frm_top, textvariable=self.in_path).grid(row=0, column=1, sticky='ew', padx=(5, 5))
        ttk.Button(frm_top, text="Обзор…", command=self.browse_in).grid(row=0, column=2)
        frm_top.columnconfigure(1, weight=1)

        # Output ZIP
        ttk.Label(frm_top, text="Выходной архив (.zip):").grid(row=1, column=0, sticky='w', pady=(5, 0))
        ttk.Entry(frm_top, textvariable=self.out_path).grid(row=1, column=1, sticky='ew', padx=(5, 5), pady=(5, 0))
        ttk.Button(frm_top, text="Сохранить как…", command=self.browse_out).grid(row=1, column=2, pady=(5, 0))

        # Rules table
        frm_rules = ttk.LabelFrame(self.root, text="Правила замены", padding=pad)
        frm_rules.pack(fill='both', expand=True, padx=pad, pady=(pad, 0))

        columns = ("mode", "pattern", "new_value")
        self.tree = ttk.Treeview(frm_rules, columns=columns, show='headings', height=10)
        self.tree.heading("mode", text="Режим")
        self.tree.heading("pattern", text="Шаблон (тег или XPath)")
        self.tree.heading("new_value", text="Новое значение")
        self.tree.column("mode", width=120, anchor='center')
        self.tree.column("pattern", width=420)
        self.tree.column("new_value", width=220)
        self.tree.pack(fill='both', expand=True, side='left')

        vsb = ttk.Scrollbar(frm_rules, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')

        # Controls under table
        frm_ctrl = ttk.Frame(self.root, padding=(pad, 0, pad, pad))
        frm_ctrl.pack(fill='x')

        self.mode_var = tk.StringVar(value='tag')
        self.pat_var = tk.StringVar()
        self.val_var = tk.StringVar()

        ttk.Label(frm_ctrl, text="Режим:").grid(row=0, column=0, sticky='w')
        ttk.Combobox(frm_ctrl, textvariable=self.mode_var, values=['tag', 'xpath'], width=7, state='readonly').grid(row=0, column=1, padx=(5, 15))
        ttk.Label(frm_ctrl, text="Шаблон:").grid(row=0, column=2)
        ttk.Entry(frm_ctrl, textvariable=self.pat_var, width=40).grid(row=0, column=3, padx=(5, 15), sticky='ew')
        ttk.Label(frm_ctrl, text="Новое значение:").grid(row=0, column=4)
        ttk.Entry(frm_ctrl, textvariable=self.val_var, width=30).grid(row=0, column=5, padx=(5, 15))
        ttk.Button(frm_ctrl, text="Добавить правило", command=self.add_rule).grid(row=0, column=6)
        ttk.Button(frm_ctrl, text="Удалить выбранные", command=self.del_selected).grid(row=0, column=7, padx=(10, 0))
        frm_ctrl.columnconfigure(3, weight=1)

        # Action row
        frm_action = ttk.Frame(self.root, padding=pad)
        frm_action.pack(fill='x')

        self.progress = ttk.Progressbar(frm_action, mode='indeterminate')
        self.progress.pack(fill='x', expand=True, side='left')

        ttk.Button(frm_action, text="Загрузить пресет…", command=self.load_preset).pack(side='left', padx=(10, 5))
        ttk.Button(frm_action, text="Сохранить пресет…", command=self.save_preset).pack(side='left')
        ttk.Button(frm_action, text="Старт", style='success.TButton' if 'ttkbootstrap' in sys.modules else 'TButton', command=self.run).pack(side='right')

        # Help
        frm_help = ttk.LabelFrame(self.root, text="Подсказка", padding=pad)
        frm_help.pack(fill='x', padx=pad, pady=(0, pad))
        ttk.Label(frm_help, text=NS_AWARE_HELP).pack(anchor='w')

    # --- Helpers ---
    def browse_in(self):
        p = filedialog.askopenfilename(title="Выберите входной архив", filetypes=[("ZIP archive", "*.zip"), ("All files", "*.*")])
        if p:
            self.in_path.set(p)
            # авто-подстановка имени результата
            if not self.out_path.get():
                self.out_path.set(self._default_out_path(p))

    def browse_out(self):
        p = filedialog.asksaveasfilename(title="Куда сохранить результат?", defaultextension=".zip", filetypes=[("ZIP archive", "*.zip")])
        if p:
            self.out_path.set(p)

    def add_rule(self):
        mode = self.mode_var.get().strip()
        pat = self.pat_var.get().strip()
        val = self.val_var.get()
        if not pat:
            messagebox.showwarning("Правило", "Укажите шаблон (тег или XPath).")
            return
        self.tree.insert('', 'end', values=(mode, pat, val))
        self.pat_var.set("")
        self.val_var.set("")

    def del_selected(self):
        for iid in self.tree.selection():
            self.tree.delete(iid)

    def rules_from_ui(self) -> List[Rule]:
        out = []
        for iid in self.tree.get_children(''):
            out.append(Rule.from_row(self.tree.item(iid, 'values')))
        return out

    def load_preset(self):
        p = filedialog.askopenfilename(title="Загрузить пресет правил", filetypes=[("JSON", "*.json"), ("All files", "*.*")])
        if not p:
            return
        try:
            with open(p, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.tree.delete(*self.tree.get_children(''))
            for it in data.get('rules', []):
                self.tree.insert('', 'end', values=(it['mode'], it['pattern'], it['new_value']))
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить пресет: {e}")

    def save_preset(self):
        p = filedialog.asksaveasfilename(title="Сохранить пресет правил", defaultextension=".json", filetypes=[("JSON", "*.json")])
        if not p:
            return
        try:
            rules = [asdict(r) for r in self.rules_from_ui()]
            with open(p, 'w', encoding='utf-8') as f:
                json.dump({"rules": rules, "version": VERSION}, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить пресет: {e}")

    def _default_out_path(self, in_path: str) -> str:
        if in_path.lower().endswith('.zip'):
            return in_path[:-4] + "_fixed.zip"
        return in_path + "_fixed.zip"

    def _maybe_load_cli_arg(self):
        # Если запущено с файлом в аргументах (перетаскивание на ярлык), подставим его как входной
        if len(sys.argv) >= 2:
            arg = sys.argv[1]
            if arg.lower().endswith('.zip'):
                self.in_path.set(arg)
                self.out_path.set(self._default_out_path(arg))

    # --- Run ---
    def run(self):
        in_zip = self.in_path.get().strip()
        out_zip = self.out_path.get().strip()
        if not in_zip:
            messagebox.showwarning("Старт", "Укажите входной архив .zip")
            return
        if not out_zip:
            messagebox.showwarning("Старт", "Укажите выходной архив .zip")
            return

        rules = self.rules_from_ui()
        if not rules:
            if not messagebox.askyesno("Правила", "Список правил пуст. Продолжить (XML будут только перепакованы)?"):
                return

        # CSV лог рядом с выходным архивом
        log_path = out_zip + ".log.csv"

        self.progress.start(10)
        self.root.configure(cursor="watch")
        self.root.update_idletasks()
        try:
            with open(log_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow(["file", "status", "details"])  # header
                stats = process_zip(in_zip, out_zip, rules, log_writer=writer)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось обработать архив: {e}")
            return
        finally:
            self.progress.stop()
            self.root.configure(cursor="")

        msg = (
            f"Готово!\n\nАрхив: {out_zip}\nЛог: {log_path}\n\n"
            f"Всего файлов: {stats.total_files}\n"
            f"XML изменено: {stats.xml_changed}\n"
            f"XML без изменений: {stats.xml_unchanged}\n"
            f"Прочие скопированы: {stats.copied_other}\n"
            f"Ошибок: {stats.errors}"
        )
        messagebox.showinfo("Результат", msg)


def main():
    app = App(TKROOT)
    TKROOT.mainloop()


if __name__ == '__main__':
    main()
