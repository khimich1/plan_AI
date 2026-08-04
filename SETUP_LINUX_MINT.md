# Настройка Linux Mint: обновления и GitHub по SSH

Инструкция для нового устройства после переноса проектов (в т.ч. через AnyDesk).

Email GitHub: `smirnov.roman.18@bk.ru`

---

## 1. Обновления системы

```bash
sudo apt update
sudo apt upgrade -y
sudo apt autoremove -y
sudo apt autoclean
```

Если Mint предлагает обновление ядра — после апгрейда:

```bash
sudo reboot
```

---

## 2. Базовые инструменты

```bash
sudo apt install -y git curl wget build-essential ca-certificates gnupg
git --version
```

---

## 3. Git: имя и email

```bash
git config --global user.name "Roman Smirnov"
git config --global user.email "smirnov.roman.18@bk.ru"
git config --global init.defaultBranch main
git config --global pull.rebase false
```

---

## 4. Подключение к GitHub по SSH

### 4.1. Создать ключ

```bash
ssh-keygen -t ed25519 -C "smirnov.roman.18@bk.ru"
```

На вопросы:

- путь к файлу — **Enter** (по умолчанию `~/.ssh/id_ed25519`)
- passphrase — по желанию (можно **Enter**, если пустая)

### 4.2. Запустить агент и добавить ключ

```bash
eval "$(ssh-agent -s)"
ssh-add ~/.ssh/id_ed25519
```

### 4.3. Скопировать публичный ключ

```bash
cat ~/.ssh/id_ed25519.pub
```

Скопируйте весь вывод (`ssh-ed25519 AAAA... smirnov.roman.18@bk.ru`).

### 4.4. Добавить ключ на GitHub

1. Откройте [https://github.com/settings/keys](https://github.com/settings/keys)
2. **New SSH key**
3. Title: например `linux-mint`
4. Key: вставьте содержимое `id_ed25519.pub`
5. **Add SSH key**

### 4.5. Проверить связь

```bash
ssh -T git@github.com
```

При первом подключении на вопрос `Are you sure you want to continue connecting` введите `yes`.

Ожидаемый ответ:

```text
Hi <username>! You've successfully authenticated...
```

---

## 5. Проверить проекты после переноса

```bash
cd ~/Code/plan_web   # или ваш путь
git status
git remote -v
```

Если remote ещё на HTTPS и нужно SSH:

```bash
git remote set-url origin git@github.com:USERNAME/REPO.git
git fetch
```

Подставьте свои `USERNAME` и `REPO` из URL репозитория на GitHub.

---

## 6. (Опционально) GitHub CLI

```bash
sudo apt install -y gh
gh auth login
```

В `gh auth login` выберите:

- GitHub.com
- SSH
- Login with a web browser (или token)

---

## 7. (Опционально) Node.js через nvm

```bash
curl -o- https://raw.githubusercontent.com/nvm-sh/nvm/v0.40.1/install.sh | bash
source ~/.bashrc
nvm install --lts
node -v && npm -v
```

---

## 8. (Опционально) Python

```bash
sudo apt install -y python3 python3-pip python3-venv
```

---

## Короткий минимум

Только обновления + Git + SSH:

```bash
sudo apt update && sudo apt upgrade -y
sudo apt install -y git curl
git config --global user.name "Roman Smirnov"
git config --global user.email "smirnov.roman.18@bk.ru"
ssh-keygen -t ed25519 -C "smirnov.roman.18@bk.ru"
eval "$(ssh-agent -s)" && ssh-add ~/.ssh/id_ed25519
cat ~/.ssh/id_ed25519.pub
# → вставить на https://github.com/settings/keys
ssh -T git@github.com
```
