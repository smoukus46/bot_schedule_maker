Гайд по запуску бота как службы

1. Создаём новый юнит systemd:

cd /lib/systemd/system
sudo nano bot.service

2. В создаваемый файл вставляем следующее содержимое:

[Unit]
Description=Bot Service
After=network.target

[Service]
Type=idle
Restart=always
RestartSec=3
User=root
WorkingDirectory=/root/Bot/
ExecStart=python3 Bot.py

[Install]
WantedBy=multi-user.target

Здесь:
Description — название службы;
WorkingDirectory — каталог, в котором содержится файл бота;
ExecStart — команда, запускающая бота;
User — учётная запись, под именем которой запускается бот;
Restart=always — указание на перезапуск службы при возникновении ошибки.

3. Закрываем файл с сохранением внесённых изменений

4. Просим systemd перечитать файлы юнитов:

sudo systemctl daemon-reload

5. Включаем новую службу:

sudo systemctl enable bot.service

6. Запускаем её:

sudo systemctl start bot.service

7. Проверяем состояние запущенной службы:

sudo systemctl status bot.service

Дополнительно:

При внесении изменений в функционал бота применяться они будут теперь при перезапуске службы:

sudo systemctl restart bot.service