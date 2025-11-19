# RPA-PROGRAMACAO-TIME-MILK RUN


pyinstaller --noconfirm --onefile --windowed --noconsole ^
 --name "RPA Milk Run" ^
 --icon "C:/Users/perna/Desktop/STALLANTIS/VIAJANTE/Viajante/Viajante.ico" ^
 --add-data "C:\Users\perna\AppData\Local\ms-playwright\chromium-1187\chrome-win;ms-playwright\chromium-1187\chrome-win" ^
 App.py
