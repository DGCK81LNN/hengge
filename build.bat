@set PATH=%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts;%PATH%
pip install -r requirements.txt
pyinstaller hengge.spec %*
@echo.
@pause
