# Плагины Полиметрики для Claude

Маркетплейс плагинов [Полиметрики](https://polimetrika.ru) — сервиса энергомониторинга производства.

## polimetrika-demo

Скилл `generate-project-pack`: Claude расспрашивает о цехе (или придумывает его сам) и собирает пакет
документов вымышленного объекта — описание, договор энергоснабжения, ведомость оборудования, 12 счетов
за электроэнергию (и, если нужно, за газ, тепло, воду и топливо), отчёты о выпуске, расшифровку разговора
с технологом и почасовые показания приборов учёта, в том числе счётчиков газа, тепла и воды на вводах. Пакет
загружается в чат нового проекта Полиметрики, помощник разбирает его и заполняет проект, а показания
приборов отправляет в хранилище показаний.

Данные согласованы между собой (одна почасовая модель объекта) и содержат заложенные потери — ответы
в `ЧТО ЗАЛОЖЕНО.md` пакета.

### Установка

**claude.ai и Claude Desktop:** скачайте `polimetrika-demo-<версия>.zip` из
[релизов](../../releases) → Customize → Plugins → Upload plugin. Нужно включённое выполнение кода.

**Claude Code:**

```bash
claude plugin marketplace add Polimetrika/claude-plugins
claude plugin install polimetrika-demo@polimetrika
```

**ChatGPT** (тарифы со скиллами — Business, Enterprise, Edu): скачайте `generate-project-pack-<версия>.zip`
из [релизов](../../releases) → Plugins → Skills → Create → Upload from your computer.

**Codex:** скопируйте папку скилла в `~/.agents/skills`:

```bash
git clone --depth 1 https://github.com/Polimetrika/claude-plugins /tmp/polimetrika-plugins
cp -r /tmp/polimetrika-plugins/plugins/polimetrika-demo/skills/generate-project-pack ~/.agents/skills/
```

Затем попросите ассистента: «Сгенерируй демо-пакет для Полиметрики».

## Разработка

```bash
python plugins/polimetrika-demo/skills/generate-project-pack/scripts/generate.py \
  plugins/polimetrika-demo/skills/generate-project-pack/examples/coffee-roastery.json --out /tmp/pack
scripts/build-zip.sh          # dist/polimetrika-demo-<версия>.zip (claude.ai) и generate-project-pack-<версия>.zip (ChatGPT)
claude plugin validate .
```

Примеры спецификаций — `examples/`: `coffee-roastery.json` (обжарочный цех, всё оборудование из каталога) и
`asphalt-plant.json` (АБЗ «Южный тракт»: 8 из 12 строк ведомости вне каталога — пилот пополнения каталога,
AI-37/docs#625 S5). Оба — регрессионный набор для каталога: после правки генератора или каталога прогоните
оба и разберите пакеты на стенде.

Генератор пишет документы под правила распознавания помощника Полиметрики (признаки типа документа,
подписи колонок). Меняя тексты и колонки, проверяйте пакет на стенде: тип каждого файла и карточки
предложений.
