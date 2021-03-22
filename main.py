from PIL import Image, ImageFilter, ImageDraw, ImageFont
import pandas
 
data = pandas.read_excel('1.xlsx', sheet_name='Лист1')
 
type = data['Вид грамоты'].tolist()
name_sur = data['ФИ участника'].tolist()
place = data['Наименование организатора мероприятия'].tolist()
reward = data['Место/награда/номинация'].tolist()
event = data['Название мероприятия'].tolist()
date = data['Дата проведения'].tolist()
 
 
headline = ImageFont.truetype('arialbd.ttf', size = 30)
 
for i in range(len(reward) - 1):
    img = Image.open('Шаблон грамоты.png')
    idraw = ImageDraw.Draw(img)
    idraw.text((300, 300), type[i], font = headline, fill='blue')
    idraw.text((300, 400), 'награждается', font = headline, fill='black')
    idraw.text((250, 500),  name_sur[i], font = headline, fill='black')
    idraw.text((250, 600), place[i], font = headline, fill='black')
    idraw.text((250, 720), reward[i], font = headline, fill='black')
    idraw.text((250, 800), event[i], font = headline, fill='black')
    idraw.text((250, 900), date[i], font = headline, fill='black')
 
    img.show()
