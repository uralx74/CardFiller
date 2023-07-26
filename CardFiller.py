from PIL import Image, ImageDraw, ImageFont, ImageColor
#from openpyxl import Workbook
import openpyxl
#import msvcrt
import os
import keyboard
import time
from os.path import abspath
from collections import namedtuple

os.system('CLS')

#gvAppPath = os.path.dirname(__file__) 
gvTemplateFileName = "bk.png"
gvDataFileName = "data.xlsx"
gvResultDir = "result"
gvResultFileExt = ".png"
#gvTemplateFileName = gvAppPath + "\\" + gvTemplateFileName
#gvDataFileName = gvAppPath + "\\" + gvDataFileName
#gvResultDir = gvAppPath + "\\" + gvResultDir
gvTemplateFileName = gvTemplateFileName
gvDataFileName = gvDataFileName
gvResultDir = gvResultDir
gvMaxRowCount = 1003

print("Файл шаблона: " + gvTemplateFileName)
print("Файл данных: " + gvDataFileName)
print("Папка для результата: " + gvResultDir)
print("Для запуска процесса нажмите Enter...")

keyboard.read_key() # Ждем нажатия клавиши

if keyboard.is_pressed('Esc'):
    print("Прервано!")
    raise SystemExit

if os.path.exists(gvResultDir) == False:
    os.mkdir(gvResultDir)

#wb = Workbook.op
#ws = wb.active
vWorkbook = openpyxl.reader.excel.load_workbook(filename=gvDataFileName,data_only=True)
vWorkbook.active = 0
vWorksheet = vWorkbook.active
  
vWsConf = vWorkbook["conf"]
vConfTextLeft = vWsConf["B2"].value
vConfTextTop = vWsConf["B3"].value
vRowInterval= vWsConf["B4"].value
vWordInterval = vWsConf["B5"].value
vCellFontName = vWsConf["B6"].value
vCellFontSize = vWsConf["B7"].value
vDefCellFontColorParam = vWsConf["B8"].value
vDefCellFontColorVal = vWsConf["B9"].value
vCreateFiles = vWsConf["B10"].value

vFont = ImageFont.truetype(vCellFontName,size=vCellFontSize)

if vCreateFiles==1:
    vImgMain = Image.open(gvTemplateFileName)

vFilesProcessed = 0
vRowsProcessed = 0
for i in range(2,gvMaxRowCount):                             # Цикл по строкам
    if keyboard.is_pressed('Esc'):
        print("Чтобы прервать процесс создания файлов нажмите Esc еще раз, чтобы продолжить нажмите Enter...")
        time.sleep(1)
        keyboard.read_key()
        if keyboard.is_pressed('Esc'):
            print("Прервано!")
            break
    
    if vWorksheet['A'+str(i)].value == None:
        break
    
    vRowsProcessed = vRowsProcessed + 1
        
    vResultFileName = vWorksheet['A'+str(i)].value + gvResultFileExt
     
    if vCreateFiles==1:
        vImgResult = vImgMain.copy()
    else:
        if os.path.exists(gvResultDir + '\\' + vResultFileName) == True:   
            vImgResult = Image.open(gvResultDir + '\\' + vResultFileName)
        else:
            print(str(vRowsProcessed) + ': ' + vResultFileName + " - файл не найден")
            continue
    
    print(str(vRowsProcessed) + ': ' + vResultFileName)
    vImageDraw2 = ImageDraw.Draw(vImgResult)
    
    vTextTop = vConfTextTop
    # Выводим текст
    for j in range(ord('B'), ord('Z') + 1):         # Цикл по столбцам 
        cell = vWorksheet[chr(j)+str(i)]
        cellH = vWorksheet[chr(j)+'1']              # Ячейка из шапки
        if cellH.value == None:                     # Если заголовок пустой, то пропускаем столбец
            continue
        
        
        #print("cell[" + chr(j)+str(i) +"]" + "color = " + str(cell.fill.start_color.index))
        #print("cell[" + chr(j)+str(i) +"]" + "font color = " + str(cell.font.color.rgb))
        
        # Параметр
        if cellH.font.color != None and type(cellH.font.color.rgb) == str:    # Если цвет текста в ячейке задан не индексом цвета
            vCellHFontColor = "#" + cellH.font.color.rgb[2:8]                 # Отбрасываем альфа-канал
        else:
            vCellHFontColor = vDefCellFontColorParam

        # Значение
        if cell.font.color != None and type(cell.font.color.rgb) == str:    # Если цвет текста в ячейке задан не индексом цвета
            vCellFontColor = "#" + cell.font.color.rgb[2:8]                 # Отбрасываем альфа-канал
        else:
            vCellFontColor = vDefCellFontColorVal
        vtext=(cellH.value).upper()
        new_box = vImageDraw2.textbbox((0,0),vtext,font=vFont)
        vTextHeader = vtext
        vTextLeft = vConfTextLeft

        if vtext[:1] != "#":                        # Если в заголовке первый символ #, то не печатаем заголовк
            vImageDraw2.text((vTextLeft,vTextTop), text=str(vTextHeader), fill=vCellHFontColor,font=vFont)
            vTextLeft = vConfTextLeft + new_box[2] + vWordInterval
     
        if cell.value != None and vtext != '''''':                              
            vImageDraw2.text((vTextLeft,vTextTop), text=str(cell.value)
                             , fill=vCellFontColor,font=vFont)  
        
        vTextTop = vTextTop + vRowInterval

    vImgResult.save(gvResultDir + '\\' + vResultFileName,'PNG')
    vImgResult.close()
    vFilesProcessed = vFilesProcessed + 1

print("------------------------------")
if vCreateFiles==1:
    print("Создано " + str(vFilesProcessed) + " файлов.")
elif vCreateFiles==0:
    print("Обработано " + str(vRowsProcessed) + " строк.")
    print("Обновлено " + str(vFilesProcessed) + " файлов.")
keyboard.read_key()
#os.system(r"explorer.exe "+gvResultDir)

#raise SystemExit


#vImageDraw.rectangle(0,0,vImgHeadWidth, vImgHeadWidth)
#vImageDraw.rounded_rectangle(((0, 0), (vImgHeadHeight,vImgHeadWidth)), 20, fill="blue")

#vResultDirFull = os.getcwd() + "\\" + gvResultDir

#vImgHead = Image.new("RGBA", (w,round(h/10)), "white")
#vImgHeadHeight, vImgHeadWidth = vImgHead.size
#vImageDraw = ImageDraw.Draw(vImgHead)

#cell_fill = sheet_active['A1'].fill.start_color.index #Получаем цвет ячейки
#cell_fill = '#' + cell_fill
#TextColor

#imgd.text((0,0), text="hello", fill="blue",font=vFont)
#img.save("img.png")

#vImageDraw2.text((60,680), text=vWorksheet['D'+str(i)].value, fill="#a000d4",font=vFont)
#vImgResult.paste(vImgHead,(0,0))
    
#img2 = Image.open("002.png")
#vImgResult = Image.blend(vImgMain, vImgHead, 0.5)
#h,w = vImgMain.size


#gvTemplateFileName  = os.path.realpath(__file__) + gvTemplateFileName
#print(">" + abspath(__file__))
#vImageDraw2.text((vTextLeft,vTextTop), text=vTextHeader, fill=ImageColor.getrgb("red"),font=vFont)