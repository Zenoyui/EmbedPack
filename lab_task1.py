#!/usr/bin/env python3
"""Задание 1: pwd, whoami, ls, ls -l, ls -a, date. Выполняются настоящие команды.
Выполнил: Степанов Глеб Владиславович, 25о9-090207-ИС/3
Запуск: python3 lab_task1.py   (из каталога practice_linux или из домашней папки)
"""
import os, subprocess

base = os.path.expanduser("~/practice_linux")
cwd = base if os.path.isdir(base) else os.getcwd()
os.makedirs(os.path.join(cwd, "results"), exist_ok=True)

out = []
for cmd in ("pwd", "whoami", "ls", "ls -l", "ls -a", "date"):
    r = subprocess.run(cmd, shell=True, cwd=cwd, capture_output=True, text=True)
    block = f"$ {cmd}\n{(r.stdout + r.stderr).rstrip()}\n"
    print(block)
    out.append(block)

path = os.path.join(cwd, "results", "task1.txt")
with open(path, "w", encoding="utf-8") as f:
    f.write("\n".join(out))
print("Сохранено:", path)
