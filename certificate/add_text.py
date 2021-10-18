import os

import pandas as pd
from PIL import Image, ImageDraw, ImageFont

font_name = ImageFont.truetype("C:Windows/Fonts/msyh.ttc", 44)
font_id = ImageFont.truetype("C:Windows/Fonts/simhei.ttf", 40)
font_content = ImageFont.truetype("C:Windows/Fonts/simhei.ttf", 30)
font_underline=ImageFont.truetype("C:Windows/Fonts/simhei.ttf", 37)

source_file_name = "小伙伴计划—2020秋开言悦读证书制作名单—张秋迎—20210306.xlsx"
template_picture = "2020秋-开言阅读证书模版-宣策组-2.jpg"
target_folder_name = "开言悦读"
ID = "2020QJ-KYYD-001"


df = pd.read_excel(source_file_name, sheet_name=0)


# 重定义尺寸
x, y = 1331, 2137

#实现字体居中功能
def pos(underline):
 return (665-len(underline)/2*37)

for index, row in df.iterrows():
    im1 = Image.open(template_picture)
    print(im1.size)
    im1 = im1.resize((x, y))
    draw = ImageDraw.Draw(im1)
    #row[1]保存用户学校
    underline=row[1]
    #row[0]保存用户姓名
    name = row[0]
    name_split = ""
    for i in name:
        name_split += i
        name_split += " "

    # 名字长度
    if len(name) == 2:
        draw.text((614, 461), name_split, (0, 0, 0), font=font_name)
    elif len(name) == 3:
        draw.text((585, 461), name_split, (0, 0, 0), font=font_name)
    elif len(name) == 4:
        draw.text((555, 461), name_split, (0, 0, 0), font=font_name)
    draw = ImageDraw.Draw(im1)

    if underline=='北京航空航天大学':
        underline='北京航空航天大学学业与发展支持中心'
    elif underline=='重庆邮电大学':
        underline='重庆邮电大学学生学业互助中心'
    elif underline=='兰州理工大学':
        underline+='学生处'
    elif underline=='中国政法大学':
        underline+='学生学业发展协会'
    elif underline=='中国人民大学':
        underline+='学生学业发展协会'
    else:
        underline=' '
    
    underline1='清华大学学生学业发展协会'
    underline2='小伙伴计划部'
    underline3='2021.03.12'

    ID_prefix = ID[:-4]
    id = f"{ID_prefix}-{str(index+1).zfill(3)}"
    draw.text((662, 350), id, (193, 211, 80), font=font_id)#金色
    
    draw.text((pos(underline),1215),underline,(0,0,0),font=font_underline)
    draw.text((pos(underline1),1215+60),underline1,(0,0,0),font=font_underline)
    draw.text((pos(underline2),1215+120),underline2,(0,0,0),font=font_underline)
    draw.text((569,1215+180),underline3,(0,0,0),font=font_underline)

    #draw.text((243, 2091), id, (255, 255, 255), font=font_id)#白色
    draw = ImageDraw.Draw(im1)
    print(name_split, id)

    # 保存

    if not os.path.exists(target_folder_name):
        os.makedirs(target_folder_name)
    im1.save(f"{target_folder_name}/{name}.jpg")

