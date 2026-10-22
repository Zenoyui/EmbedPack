#!/bin/bash
# Lab: Bash/Linux, tasks 1-7 + final
# Run: bash run_steps.sh (terminal maximized)

read -rp "Фамилия И.О.: " STUDENT
read -rp "Группа: " GROUP
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
WORK="$HOME/practice_linux"

# old dir -> practice_linux_old_HHMMSS
if [ -e "$WORK" ]; then mv "$WORK" "${WORK}_old_$(date +%H%M%S)"; fi

SHOTS="$HERE/screenshots"; mkdir -p "$SHOTS"; N=0
shot() {  # screen grab
  N=$((N+1)); f="$SHOTS/$(printf '%02d' $N)_$1.png"; sleep 1
  if   command -v gnome-screenshot >/dev/null; then gnome-screenshot -w -f "$f" 2>/dev/null
  elif [ -n "$WAYLAND_DISPLAY" ] && python3 "$HERE/portal_shot.py" "$f" 2>/dev/null; then :
  elif command -v scrot >/dev/null;            then scrot -u "$f" 2>/dev/null
  elif command -v import >/dev/null;           then import -window root "$f" 2>/dev/null
  elif [[ "$XDG_CURRENT_DESKTOP" == *KDE* ]] && command -v spectacle >/dev/null; then spectacle -a -b -n -o "$f" 2>/dev/null
  else echo "shot: no tool (gnome-screenshot, scrot, import), take manually (Print Screen), then Enter"; read -rp "" _; return; fi
  [ -s "$f" ] || { echo "shot: file empty (path=$f), take manually (Print Screen), then Enter"; read -rp "" _; }
}
CUR=""
printf '\e[8;45;110t' 2>/dev/null   # resize
prompt() { printf '%s@%s:%s$ ' "$(whoami)" "$(hostname)" "${PWD/#$HOME/\~}"; }
step() {  # next screen
  if [ -n "$CUR" ]; then shot "$CUR"; fi
  CUR="$2"; clear
  echo "# $1 | $STUDENT, группа $GROUP | $(date '+%d.%m.%Y %H:%M:%S')"
}
c() { prompt; echo "$1"; eval "$1"; }      # one cmd
m() {                                      # multiline
  local first=1 line
  while IFS= read -r line; do
    if [ $first = 1 ]; then prompt; first=0; else printf '> '; fi
    echo "$line"
  done <<< "$1"
  eval "$1"
}

cd "$HOME" || exit 1

step "Задание 1. Основные команды" task1
c "mkdir practice_linux"
c "cd practice_linux"
c "pwd"; c "whoami"; c "ls"; c "ls -l"; c "ls -a"; c "date"

step "Задание 2. Создание каталогов и файлов" task2
c "mkdir data scripts results"
c "ls -l"
c "touch students.txt"
m "cat > students.txt << 'EOF'
Иванов И.И. 2005 Москва 4.5
Петров П.С. 2006 Казань 4.2
Сидорова А.В. 2005 Омск 4.8
Смирнов Д.А. 2006 Томск 3.9
Кузнецова М.И. 2005 Самара 4.6
EOF"
c "cat students.txt"
c "echo 'тестовая строка' > test.txt"
c "cat test.txt"
c "mv test.txt example.txt"
c "ls"

step "Задание 3. Обработка текстовых данных" task3
c "wc -l students.txt"
c 'grep "Иванов" students.txt'
c 'grep "Казань" students.txt'
c "sort students.txt"
c "sort students.txt > sorted_students.txt"
c "cat sorted_students.txt"
c "wc -l students.txt"
c "wc -l sorted_students.txt"

step "Задание 4. Переменные" task4
c 'student="Иванов И.И."'
c 'year=2005'
c 'city="Москва"'
c 'echo "$student"'; c 'echo "$year"'; c 'echo "$city"'
c 'filename="students.txt"'
c 'cat "$filename"'
c 'count=$(wc -l < "$filename")'
c 'echo "Количество записей: $count"'
c 'name="Иванов И.И."; grade=4.5'
c 'echo "Студент: $name; город: $city; средний балл: $grade"'
c 'city="Казань"'
c 'echo "Студент: $name; город: $city; средний балл: $grade"'

step "Задание 5. Условные конструкции" task5
for g in 4 3; do
  c "grade=$g"
  m 'if [ "$grade" -ge 4 ]; then
    echo "Условие выполнено"
else
    echo "Условие не выполнено"
fi'
done
for f in students.txt missing.txt; do
  m "if [ -f \"$f\" ]; then echo \"Файл найден\"; else echo \"Файл не найден\"; fi"
done
for n in 5 3 7; do
  c "count=$n"
  m 'if [ "$count" -ge 5 ]; then echo "Записей достаточно"; else echo "Необходимо добавить записи"; fi'
done

step "Задание 6. Циклы" task6
m 'for i in 1 2 3 4 5 6 7 8 9 10; do echo "Запись № $i"; done'
m 'i=1
while [ "$i" -le 7 ]; do echo "Номер: $i"; i=$((i + 1)); done'
m 'while IFS= read -r line; do echo "$line"; done < students.txt'

step "Задание 7. Передача параметров сценарию" task7
c "cd scripts"
m "cat > parameters.sh << 'EOF'
#!/bin/bash
echo \"Имя сценария: \$0\"
echo \"Фамилия и инициалы: \$1\"
echo \"Город: \$2\"
echo \"Год рождения: \$3\"
echo \"Количество параметров: \$#\"
EOF"
c 'bash parameters.sh "Иванов И.И." "Москва" 2005'
c 'bash parameters.sh "Петров П.С." "Казань" 2006'
c "cd .."

step "Итоговое задание. Обработка students.txt" final
m "cat > scripts/check_students.sh << 'EOF'
#!/bin/bash
file=\"\${1:-students.txt}\"
if [ -f \"\$file\" ]; then
    cat \"\$file\"
    count=\$(wc -l < \"\$file\")
    echo \"Количество записей: \$count\"
    if [ \"\$count\" -ge 5 ]; then echo \"Записей достаточно\"; else echo \"Необходимо добавить записи\"; fi
else
    echo \"file not found (path=\$file)\"
fi
EOF"
c "bash scripts/check_students.sh students.txt"
c "bash scripts/check_students.sh missing.txt"

shot "$CUR"; echo; echo "screenshots: $SHOTS"
