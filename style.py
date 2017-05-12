"""
Стили оформления ячеек для телефонной книги
"""
from xlwt import *

font0 = Font()
font0.name = 'Arial'
font0.bold = True

font1 = Font()
font1.name = 'Arial'
font1.italic = True

al0 = Alignment()
al0.horz = Alignment.HORZ_CENTER

# Заголовок
bh0 = Borders()
bh0.left = 2
bh0.right = 1
bh0.top = 2
bh0.bottom = 1
bh0.bottom_colour = Style.colour_map['gray25']
bh0.left_colour = Style.colour_map['gray25']
bh0.right_colour = Style.colour_map['gray25']
bh0.top_colour = Style.colour_map['gray25']

sh0 = XFStyle()
sh0.borders = bh0

bh1 = Borders()
bh1.left = 1
bh1.right = 7
bh1.top = 2
bh1.bottom = 1
bh1.bottom_colour = Style.colour_map['gray25']
bh1.left_colour = Style.colour_map['gray25']
bh1.right_colour = Style.colour_map['gray25']
bh1.top_colour = Style.colour_map['gray25']

sh1 = XFStyle()
sh1.alignment = al0
sh1.borders = bh1
sh1.font = font0

bh2 = Borders()
bh2.left = 7
bh2.right = 7
bh2.top = 2
bh2.bottom = 1
bh2.bottom_colour = Style.colour_map['gray25']
bh2.left_colour = Style.colour_map['gray25']
bh2.right_colour = Style.colour_map['gray25']
bh2.top_colour = Style.colour_map['gray25']

sh2 = XFStyle()
sh2.alignment = al0
sh2.borders = bh2
sh2.font = font0

bh3 = Borders()
bh3.left = 7
bh3.right = 2
bh3.top = 2
bh3.bottom = 1
bh3.bottom_colour = Style.colour_map['gray25']
bh3.left_colour = Style.colour_map['gray25']
bh3.right_colour = Style.colour_map['gray25']
bh3.top_colour = Style.colour_map['gray25']

sh3 = XFStyle()
sh3.alignment = al0
sh3.borders = bh3
sh3.font = font0

# Организация
bo0 = Borders()
bo0.left = 2
bo0.right = 2
bo0.top = 1
bo0.bottom = 1
bo0.bottom_colour = Style.colour_map['gray25']
bo0.left_colour = Style.colour_map['gray25']
bo0.right_colour = Style.colour_map['gray25']
bo0.top_colour = Style.colour_map['gray25']

so0 = XFStyle()
so0.borders = bo0
so0.font = font0

# Отдел
bd0 = Borders()
bd0.left = 2
bd0.right = 1
bd0.top = 3
bd0.bottom = 3
bd0.bottom_colour = Style.colour_map['gray25']
bd0.left_colour = Style.colour_map['gray25']
bd0.right_colour = Style.colour_map['gray25']
bd0.top_colour = Style.colour_map['gray25']

sd0 = XFStyle()
sd0.borders = bd0

bd1 = Borders()
bd1.left = 1
bd1.right = 2
bd1.top = 3
bd1.bottom = 3
bd1.bottom_colour = Style.colour_map['gray25']
bd1.left_colour = Style.colour_map['gray25']
bd1.right_colour = Style.colour_map['gray25']
bd1.top_colour = Style.colour_map['gray25']

sd1 = XFStyle()
sd1.borders = bd1
sd1.font = font1

# Сотрудники
bs0 = Borders()
bs0.left = 2
bs0.right = 1
bs0.top = 7
bs0.bottom = 7
bs0.bottom_colour = Style.colour_map['gray25']
bs0.left_colour = Style.colour_map['gray25']
bs0.right_colour = Style.colour_map['gray25']
bs0.top_colour = Style.colour_map['gray25']

ss0 = XFStyle()
ss0.borders = bs0

bs1 = Borders()
bs1.left = 3
bs1.right = 7
bs1.top = 7
bs1.bottom = 7
bs1.bottom_colour = Style.colour_map['gray25']
bs1.left_colour = Style.colour_map['gray25']
bs1.right_colour = Style.colour_map['gray25']
bs1.top_colour = Style.colour_map['gray25']

ss1 = XFStyle()
ss1.borders = bs1

bs2 = Borders()
bs2.left = 7
bs2.right = 7
bs2.top = 7
bs2.bottom = 7
bs2.bottom_colour = Style.colour_map['gray25']
bs2.left_colour = Style.colour_map['gray25']
bs2.right_colour = Style.colour_map['gray25']
bs2.top_colour = Style.colour_map['gray25']

ss2 = XFStyle()
ss2.borders = bs2

bs3 = Borders()
bs3.left = 7
bs3.right = 2
bs3.top = 7
bs3.bottom = 7
bs3.bottom_colour = Style.colour_map['gray25']
bs3.left_colour = Style.colour_map['gray25']
bs3.right_colour = Style.colour_map['gray25']
bs3.top_colour = Style.colour_map['gray25']

ss3 = XFStyle()
ss3.borders = bs3

# Подвал
bb0 = Borders()
bb0.top = 2
bb0.top_colour = Style.colour_map['gray25']

sb0 = XFStyle()
sb0.borders = bb0
