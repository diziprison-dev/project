@echo off
echo ========================================
echo  Urun Fiyat Hesaplayici - Kurulum
echo ========================================
echo.

echo Python'un yuklu oldugunu kontrol ediliyor...
python --version >nul 2>&1
if errorlevel 1 (
    echo HATA: Python yuklu degil!
    echo.
    echo Lutfen Python'u yukleyin:
    echo https://www.python.org/downloads/
    echo.
    echo Yukleme sirasinda "Add Python to PATH" secenegini isaretleyin!
    echo.
    pause
    exit /b 1
)

echo Python bulundu!
python --version
echo.

echo Gerekli paketler yukleniyor...
echo.

echo [1/3] pandas yukleniyor...
pip install pandas>=1.5.0
if errorlevel 1 (
    echo HATA: pandas yuklenemedi!
    pause
    exit /b 1
)

echo [2/3] openpyxl yukleniyor...
pip install openpyxl>=3.0.0
if errorlevel 1 (
    echo HATA: openpyxl yuklenemedi!
    pause
    exit /b 1
)

echo [3/3] pyinstaller yukleniyor...
pip install pyinstaller>=5.0.0
if errorlevel 1 (
    echo HATA: pyinstaller yuklenemedi!
    pause
    exit /b 1
)

echo.
echo ========================================
echo  Kurulum Basarili!
echo ========================================
echo.
echo Artik programi calistirabilirsiniz:
echo - run_app.bat dosyasini cift tiklayin
echo - veya: python main.py
echo.
echo Windows executable olusturmak icin:
echo - build_advanced.py dosyasini calistirin
echo.
echo NOT: sturmmm.xlsx dosyasi program ile ayni klasorde olmalidir!
echo.
pause
