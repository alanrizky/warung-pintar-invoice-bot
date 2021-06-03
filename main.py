import glob
import locale
import os.path
import re
import shutil
import subprocess
import time
from pathlib import *

import pyautogui
import pygetwindow as gw
import pywinauto
from pywinauto.keyboard import send_keys

import rename_process

try:
    for new_txt in glob.glob("*.txt"):
        os.remove(new_txt)
except FileNotFoundError:
    pass

for filename in glob.glob(os.path.join(r"D:\Download\Compressed\Compressed", '*.*')):
    shutil.copy(filename, "Text File")

rename_process.rename()

dir = "Text File\*.txt"
move_dir = "Text File"
file = glob.glob(dir)
locale.setlocale(locale.LC_ALL, "id_ID")
current_time = time.strftime("%A, %d %B %Y")

sum_files = str(len([x for x in os.listdir(os.path.dirname(dir))]))

if sum_files == "0":
    print("Gak ada invoice yang diproses. Program akan exit.")
    exit()
else:
    print("Eksekusi program dimulai pada " + time.strftime("%X\n"))
    print("Ada " + sum_files + " invoice yang diproses")

tanggal = str(time.strftime("%d"))
convert_tanggal = int(tanggal)
bulan = str(time.strftime("%#m"))
tahun = str(time.strftime("%Y"))

excel = 'C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE'
excel_jeje = 'D:\\Download\\CHROME\\JEJE\\DAILY REPORT MANUAL OUTBOUND 2021.xlsx'

subprocess.Popen([excel, excel_jeje])

pattern_item = re.compile(
    "(?<=\] ).*|\d+(,\d{1,5})? (Bal|Batang|Botol|Bungkus|Dus|Kaleng|Kotak|Pak|Pcs|Pouch|Renceng|Slop|Pack|Sachet|Box"
    "|Gelas)|(?<=\) ).*|WP\d{11,15}\s(.*)\n|SO\d{6,9}")

for file_name in file:
    with open(os.path.join(r"New", os.path.basename(file_name)), "w") as title:
        title.write(os.path.basename(file_name) + "\n")
    f = open(file_name, 'r+')
    lst = []
    for x in f:
        lst.append(x)
    f.close()
    for i, line in enumerate(open(file_name)):
        with open(os.path.join(r"New", os.path.basename(file_name)), "a+") as f:
            for match in re.finditer(pattern_item, line):
                f.write(match.group() + "\n")

    with open(os.path.join(r"New", os.path.basename(file_name)), 'r+') as a:
        text = a.read()
        lines = text.split("\n")
        not_empty = [line for line in lines if line.strip() != ""]
        x = ""
        for line in not_empty:
            x += line + "\n"
        a.seek(0)
        a.write(re.sub(r"\n([\d])", r"\1", x))
        a.truncate()
        a.close()

finished_file = glob.glob(r"New\*.txt")

with open(str(current_time) + ".txt", "wb") as out:
    for x in finished_file:
        with open(x, "rb") as i:
            out.write(i.read())

vs_code = 'C:\\Users\\ALAN\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe'
file_jadi = 'D:\\PENTING\\SOFTWARE\\Program mbuat sendiri\\wp jeje\\{}.txt'.format(str(current_time))
subprocess.Popen([vs_code, file_jadi])

for new_txt in glob.glob("New\\*.txt"):
    os.remove(new_txt)


def focus_to_window(window_title=None):
    window = gw.getWindowsWithTitle(window_title)[0]
    if not window.isActive:
        pywinauto.application.Application().connect(handle=window._hWnd).top_window().set_focus()


def maximize_window(window_title=None):
    window = gw.getWindowsWithTitle(window_title)[0]
    if not window.isMaximized:
        window.maximize()


time.sleep(2)

focus_to_window("Code")
maximize_window("Code")
send_keys('^f')
pyautogui.write(str("^(?!.*.\.[A-Z]|WP\d{11,15}\s(.*)\n|SO\d{6}).*$"))
send_keys('{TAB}{TAB}{TAB}')
send_keys('{SPACE}')
send_keys('^+l')
send_keys('{ESC}{ESC}')
send_keys('^{LEFT}^{LEFT}')
pyautogui.write(str("~"))
send_keys('^{RIGHT}^{RIGHT}')
send_keys('^{LEFT}')
pyautogui.write(str("~"))
send_keys('^f')
pyautogui.write(str("^(?!.*.\.[A-Z]|WP\d{11,15}\s(.*)\n|SO\d{6}).*$"))
send_keys('^+l')
send_keys('{ESC}')
send_keys('^c')

focus_to_window("Excel")
maximize_window("Excel")
send_keys('{TAB}', pause=0.01)
pyautogui.write(bulan + "/{}".format(convert_tanggal) + "/" + tahun)
send_keys('{TAB}', pause=0.01)
pyautogui.write(str('B 9576 FCE'))
send_keys('{TAB}{TAB}{TAB}', pause=0.01)
send_keys('%hvun%s%o')
pyautogui.write(str("~"))
send_keys('%f', pause=0.01)

focus_to_window("Code")

send_keys('{BACKSPACE}')
send_keys('^f')
pyautogui.write(str("WP\d{11,15}\s(.*)\n"))
send_keys('^+l')
send_keys('{ESC}')
send_keys('{RIGHT}')
send_keys('{DELETE}')
pyautogui.write(str(","))
send_keys('{END}')
send_keys('{DELETE}')
send_keys('^f')
pyautogui.write(str('^(?!.*WP\d{11,15}\s(.*)\n|SO\d{6}|\\n).*$'))
send_keys('^+l')
send_keys('{ESC}')
send_keys('{BACKSPACE}')
send_keys('{DELETE}')
send_keys('^a')
send_keys('^c')
focus_to_window("Excel")
send_keys('{LEFT}{LEFT}', pause=0.01)
send_keys('%hvun%c%sf')
send_keys('{ENTER}', pause=0.01)

focus_to_window("Code")
send_keys('{ESC}')
send_keys('^z^z')
send_keys('^f')
pyautogui.write(str('^(?!.*WP\d{11,15}\s(.*)\n|SO\d{6}|\\n).*$'))
send_keys('+{ENTER}+{ENTER}')
pyautogui.hotkey('winleft', 'left')

# FCE
focus_to_window("Excel")
pyautogui.hotkey('winleft', 'right')
send_keys('{LEFT}', pause=0.01)
send_keys('{UP}', pause=0.01)
send_keys('^{RIGHT}', pause=0.01)
send_keys('{RIGHT}{RIGHT}{RIGHT}{RIGHT}', pause=0.01)
pyautogui.write(str('Dyan Kurniawan'))
send_keys('{RIGHT}', pause=0.01)
pyautogui.write(str('Bagus Santoso'))

# NBA
send_keys('{HOME}', pause=0.01)
send_keys('{RIGHT}{RIGHT}', pause=0.01)
focus_to_window("Code")
if os.path.exists(r"Text File/B.FCE.a.txt"):
    send_keys('{ENTER}{ENTER}', pause=0.01)
else:
    send_keys('{ENTER}', pause=0.01)
send_keys('{ESC}', pause=0.01)
send_keys('{DOWN}', pause=0.01)
send_keys('{END}', pause=0.01)
send_keys('^+{LEFT}', pause=0.01)
send_keys('^c', pause=0.01)
focus_to_window("Excel")
send_keys('^f', pause=0.01)
send_keys('^v', pause=0.01)
send_keys('{ENTER}', pause=0.01)
send_keys('{ESC}', pause=0.01)
send_keys('{LEFT}{LEFT}', pause=0.01)
pyautogui.write(str('N 8764 BA'))
send_keys('^{RIGHT}', pause=0.01)
send_keys('{RIGHT}{RIGHT}{RIGHT}{RIGHT}', pause=0.01)
pyautogui.write(str('Baskara Cahya Tama'))
send_keys('{RIGHT}', pause=0.01)
pyautogui.write(str('Mochamad Soni'))

# LVC
focus_to_window("Code")
send_keys('^f', pause=0.01)
send_keys('^a', pause=0.01)
pyautogui.write(str('LVC'))
send_keys('{ESC}', pause=0.01)
send_keys('{DOWN}', pause=0.01)
send_keys('{END}', pause=0.01)
send_keys('^+{LEFT}', pause=0.01)
send_keys('^c', pause=0.01)
focus_to_window("Excel")
send_keys('^f', pause=0.01)
send_keys('^v', pause=0.01)
send_keys('{ENTER}', pause=0.01)
send_keys('{ESC}', pause=0.01)
send_keys('{LEFT}{LEFT}', pause=0.01)
pyautogui.write(str('L 8678 VC'))
send_keys('^{RIGHT}', pause=0.01)
send_keys('{RIGHT}{RIGHT}{RIGHT}{RIGHT}', pause=0.01)
pyautogui.write(str('Tangguh Novariani'))
send_keys('{RIGHT}', pause=0.01)
pyautogui.write(str('Oktavian Perwira Dirgantara'))

# LBU
focus_to_window("Code")
send_keys('^f', pause=0.01)
send_keys('^a', pause=0.01)
pyautogui.write(str('LBU'))
send_keys('{ESC}', pause=0.01)
send_keys('{DOWN}', pause=0.01)
send_keys('{END}', pause=0.01)
send_keys('^+{LEFT}', pause=0.01)
send_keys('^c', pause=0.01)
focus_to_window("Excel")
send_keys('^f', pause=0.01)
send_keys('^v', pause=0.01)
send_keys('{ENTER}', pause=0.01)
send_keys('{ESC}', pause=0.01)
send_keys('{LEFT}{LEFT}', pause=0.01)
pyautogui.write(str('L 9752 BU'))
send_keys('^{RIGHT}', pause=0.01)
send_keys('{RIGHT}{RIGHT}{RIGHT}{RIGHT}', pause=0.01)
pyautogui.write(str('Rizki Prayoga'))
send_keys('{RIGHT}', pause=0.01)
pyautogui.write(str('Moch Diky'))

# NBC
focus_to_window("Code")
send_keys('^f', pause=0.01)
send_keys('^a', pause=0.01)
pyautogui.write(str('NBC'))
send_keys('{ESC}', pause=0.01)
send_keys('{DOWN}', pause=0.01)
send_keys('{END}', pause=0.01)
send_keys('^+{LEFT}', pause=0.01)
send_keys('^c', pause=0.01)
focus_to_window("Excel")
send_keys('^f', pause=0.01)
send_keys('^v', pause=0.01)
send_keys('{ENTER}', pause=0.01)
send_keys('{ESC}', pause=0.01)
send_keys('{LEFT}{LEFT}', pause=0.01)
pyautogui.write(str('N 8108 BC'))
send_keys('^{RIGHT}', pause=0.01)
send_keys('{RIGHT}{RIGHT}{RIGHT}{RIGHT}', pause=0.01)
pyautogui.write(str('Dwi Indra'))
send_keys('{RIGHT}', pause=0.01)
pyautogui.write(str('Dwi Indra H'))

# MLG01
if os.path.exists(r"Text File/K.MLG1.txt"):
    focus_to_window("Code")
    send_keys('^f', pause=0.01)
    send_keys('^a', pause=0.01)
    pyautogui.write(str('MLG1'))
    send_keys('{ESC}', pause=0.01)
    send_keys('{DOWN}', pause=0.01)
    send_keys('{END}', pause=0.01)
    send_keys('^+{LEFT}', pause=0.01)
    send_keys('^c', pause=0.01)
    focus_to_window("Excel")
    send_keys('^f', pause=0.01)
    send_keys('^v', pause=0.01)
    send_keys('{ENTER}', pause=0.01)
    send_keys('{ESC}', pause=0.01)
    send_keys('{LEFT}{LEFT}', pause=0.01)
    pyautogui.write(str('MLG01'))
    send_keys('^{RIGHT}', pause=0.01)
    send_keys('{RIGHT}{RIGHT}{RIGHT}{RIGHT}', pause=0.01)
    pyautogui.write(str('3PL_01_MLG'))
    send_keys('{RIGHT}', pause=0.01)
    pyautogui.write(str('Jaya Sadewa'))
    send_keys('{HOME}', pause=0.01)

# MLG02
if os.path.exists(r"Text File/M.MLG2.txt"):
    focus_to_window("Code")
    send_keys('^f', pause=0.01)
    send_keys('^a', pause=0.01)
    pyautogui.write(str('MLG2'))
    send_keys('{ESC}', pause=0.01)
    send_keys('{DOWN}', pause=0.01)
    send_keys('{END}', pause=0.01)
    send_keys('^+{LEFT}', pause=0.01)
    send_keys('^c', pause=0.01)
    focus_to_window("Excel")
    send_keys('^f', pause=0.01)
    send_keys('^v', pause=0.01)
    send_keys('{ENTER}', pause=0.01)
    send_keys('{ESC}', pause=0.01)
    send_keys('{LEFT}{LEFT}', pause=0.01)
    pyautogui.write(str('MLG02'))
    send_keys('^{RIGHT}', pause=0.01)
    send_keys('{RIGHT}{RIGHT}{RIGHT}{RIGHT}', pause=0.01)
    pyautogui.write(str('3PL_02_MLG'))
    send_keys('{RIGHT}', pause=0.01)
    pyautogui.write(str('Okta Saputra'))
    send_keys('{HOME}', pause=0.01)

# MLG03
if os.path.exists(r"Text File/O.MLG3.txt"):
    focus_to_window("Code")
    send_keys('^f', pause=0.01)
    send_keys('^a', pause=0.01)
    pyautogui.write(str('MLG3'))
    send_keys('{ESC}', pause=0.01)
    send_keys('{DOWN}', pause=0.01)
    send_keys('{END}', pause=0.01)
    send_keys('^+{LEFT}', pause=0.01)
    send_keys('^c', pause=0.01)
    focus_to_window("Excel")
    send_keys('^f', pause=0.01)
    send_keys('^v', pause=0.01)
    send_keys('{ENTER}', pause=0.01)
    send_keys('{ESC}', pause=0.01)
    send_keys('{LEFT}{LEFT}', pause=0.01)
    pyautogui.write(str('MLG03'))
    send_keys('^{RIGHT}', pause=0.01)
    send_keys('{RIGHT}{RIGHT}{RIGHT}{RIGHT}', pause=0.01)
    pyautogui.write(str('-----'))
    send_keys('{RIGHT}', pause=0.01)
    pyautogui.write(str('------'))
    send_keys('{HOME}', pause=0.01)

# MLG06
if os.path.exists(r"Text File/Q.MLG6.txt"):
    focus_to_window("Code")
    send_keys('^f', pause=0.01)
    send_keys('^a', pause=0.01)
    pyautogui.write(str('MLG06'))
    send_keys('{ESC}', pause=0.01)
    send_keys('{DOWN}', pause=0.01)
    send_keys('{END}', pause=0.01)
    send_keys('^+{LEFT}', pause=0.01)
    send_keys('^c', pause=0.01)
    focus_to_window("Excel")
    send_keys('^f', pause=0.01)
    send_keys('^v', pause=0.01)
    send_keys('{ENTER}', pause=0.01)
    send_keys('{ESC}', pause=0.01)
    send_keys('{LEFT}{LEFT}', pause=0.01)
    pyautogui.write(str('MLG06'))
    send_keys('^{RIGHT}', pause=0.01)
    send_keys('{RIGHT}{RIGHT}{RIGHT}{RIGHT}', pause=0.01)
    pyautogui.write(str('3PL_01_MLG'))
    send_keys('{RIGHT}', pause=0.01)
    pyautogui.write(str('Okta Saputra'))
    send_keys('{HOME}', pause=0.01)

# MLG07
if os.path.exists(r"Text File/R.MLG7.txt"):
    focus_to_window("Code")
    send_keys('^f', pause=0.01)
    send_keys('^a', pause=0.01)
    pyautogui.write(str('MLG07'))
    send_keys('{ESC}', pause=0.01)
    send_keys('{DOWN}', pause=0.01)
    send_keys('{END}', pause=0.01)
    send_keys('^+{LEFT}', pause=0.01)
    send_keys('^c', pause=0.01)
    focus_to_window("Excel")
    send_keys('^f', pause=0.01)
    send_keys('^v', pause=0.01)
    send_keys('{ENTER}', pause=0.01)
    send_keys('{ESC}', pause=0.01)
    send_keys('{LEFT}{LEFT}', pause=0.01)
    pyautogui.write(str('MLG07'))
    send_keys('^{RIGHT}', pause=0.01)
    send_keys('{RIGHT}{RIGHT}{RIGHT}{RIGHT}', pause=0.01)
    pyautogui.write(str('3PL_07_MLG'))
    send_keys('{RIGHT}', pause=0.01)
    pyautogui.write(str('Oktavian Perwira Dirgantara'))
    send_keys('{HOME}', pause=0.01)

# font
send_keys('{ENTER}', pause=0.01)
send_keys('{HOME}', pause=0.01)
focus_to_window("Excel")
send_keys('^a%hffcalibri{ENTER}', pause=0.01)
send_keys('^a^{UP}', pause=0.01)

for new_txt in glob.glob("Text File\\*.txt"):
    os.remove(new_txt)

# Open all invoices
dir = Path(r"D:\Download\CHROME")
pdf_files = list(dir.glob('*.pdf'))
for pdfs in pdf_files:
    os.startfile(str(pdfs))

time.sleep(1.7)

# Find non angka
focus_to_window("Excel")
maximize_window("Excel")
send_keys('{DOWN}', pause=0.01)
send_keys('^{RIGHT}^{RIGHT}{LEFT}', pause=0.01)
send_keys('^+{DOWN}%hso{DOWN}{ENTER}', pause=0.01)

# end task Code.exe
time.sleep(0.1)
focus_to_window("Code")
send_keys('^w', pause=0.01)
send_keys('{TAB}', pause=0.01)
send_keys('{SPACE}', pause=0.01)
send_keys('%{F4}', pause=0.01)

print("\nBerakhir pada " + time.strftime("%X"))
