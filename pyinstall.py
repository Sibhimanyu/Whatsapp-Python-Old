# https://pyinstaller.org/en/stable/usage.html

import pyinstaller._WhatsappSender_

pyinstaller._WhatsappSender_.run([
    'main.py',
    '--windowed',
    '--noconsole',
    '--icon=auro.icns'
])

"""pyinstaller --windowed --onefile --noconsole --icon=auro.ico --add-data="red.json:." --add-data="whatsapp.json:." --add-data="parameters.json:." --add-data="auro.ico:." WhatsappSender.py """