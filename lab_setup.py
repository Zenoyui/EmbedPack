#!/usr/bin/env python3
"""Лабораторная работа: Bash/Linux. Выполнил Степанов Глеб Владиславович, 25о9-090207-ИС/3.
Создаёт каталоги и файлы из работы, прогоняет команды и пишет протокол protocol.txt.
Запуск: python3 lab_setup.py   (на Windows: python lab_setup.py)
"""
import os, shutil, subprocess, datetime

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "practice_linux")
STUDENTS = """Иванов И.И. 2005 Москва 4.5
Петров П.С. 2006 Казань 4.2
Сидорова А.В. 2005 Омск 4.8
Смирнов Д.А. 2006 Томск 3.9
Кузнецова М.И. 2005 Самара 4.6
"""
PARAMETERS_SH = """#!/bin/bash
echo "Имя сценария: $0"
echo "Фамилия и инициалы: $1"
echo "Город: $2"
echo "Год рождения: $3"
echo "Количество параметров: $#"
"""
CHECK_SH = """#!/bin/bash
# Итоговое задание: обработка students.txt
file="${1:-students.txt}"
if [ -f "$file" ]; then
    cat "$file"
    count=$(wc -l < "$file")
    echo "Количество записей: $count"
    if [ "$count" -ge 5 ]; then
        echo "Записей достаточно"
    else
        echo "Необходимо добавить записи"
    fi
else
    echo "Ошибка: файл $file не найден"
fi
"""
log = []
def say(t=""):
    print(t); log.append(t)

def wr(path, text):
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)

if os.path.exists(ROOT):
    shutil.rmtree(ROOT)
say(f"Протокол выполнения. {datetime.datetime.now():%d.%m.%Y %H:%M:%S}")
say("Выполнил: Степанов Глеб Владиславович, группа 25о9-090207-ИС/3\n")

say("$ mkdir practice_linux && cd practice_linux && mkdir data scripts results")
os.makedirs(ROOT)
for d in ("data", "scripts", "results"):
    os.mkdir(os.path.join(ROOT, d))
say("\n".join(sorted(os.listdir(ROOT))))

say("\n$ touch students.txt ; nano students.txt ; cat students.txt")
wr(os.path.join(ROOT, "students.txt"), STUDENTS)
say(STUDENTS.rstrip())

lines = STUDENTS.splitlines()
say("\n$ wc -l students.txt")
say(f"{len(lines)} students.txt")
say('\n$ grep "Иванов" students.txt')
say("\n".join(l for l in lines if "Иванов" in l))
say('$ grep "Казань" students.txt')
say("\n".join(l for l in lines if "Казань" in l))
say("\n$ sort students.txt > sorted_students.txt ; cat sorted_students.txt")
srt = sorted(lines)
wr(os.path.join(ROOT, "sorted_students.txt"), "\n".join(srt) + "\n")
say("\n".join(srt))
say(f"$ wc -l sorted_students.txt\n{len(srt)} sorted_students.txt")

wr(os.path.join(ROOT, "scripts", "parameters.sh"), PARAMETERS_SH)
wr(os.path.join(ROOT, "scripts", "check_students.sh"), CHECK_SH)
for s in ("parameters.sh", "check_students.sh"):
    try: os.chmod(os.path.join(ROOT, "scripts", s), 0o755)
    except OSError: pass
say("\nСозданы сценарии: scripts/parameters.sh, scripts/check_students.sh")

bash = shutil.which("bash")
if bash:
    def run(args, cwd):
        r = subprocess.run([bash] + args, cwd=cwd, capture_output=True, text=True, encoding="utf-8")
        say("$ bash " + " ".join(f'"{a}"' if " " in a else a for a in args))
        say((r.stdout + r.stderr).rstrip())
    sc = os.path.join(ROOT, "scripts")
    say("")
    run(["parameters.sh", "Иванов И.И.", "Москва", "2005"], sc)
    run(["parameters.sh", "Петров П.С.", "Казань", "2006"], sc)
    say("")
    run(["check_students.sh", os.path.join(ROOT, "students.txt")], sc)
    run(["check_students.sh", "missing.txt"], sc)
else:
    say("\nbash не найден на этом компьютере: сценарии созданы, но не запущены.")
    say("Запустите их в Linux/WSL/Git Bash: bash scripts/parameters.sh \"Иванов И.И.\" \"Москва\" 2005")

wr(os.path.join(ROOT, "results", "protocol.txt"), "\n".join(log) + "\n")
print(f"\nГотово. Папка: {ROOT}\nПротокол: {os.path.join(ROOT, 'results', 'protocol.txt')}")
