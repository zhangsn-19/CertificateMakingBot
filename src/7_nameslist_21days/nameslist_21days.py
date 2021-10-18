
"""
用于每日一事活动。从小打卡后台导出的数据文件直接生成可以用的获奖名单。
2020年10月21日
"""
import datetime
import logging

import pandas as pd
import xlwings as xw

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
    # filename="record.txt",
)
logger = logging.getLogger(__name__)
date_str = datetime.datetime.now().strftime("%Y%m%d")


# ************************************************************************
# 在这里修改文件名
# ************************************************************************

activity_name = "2020秋晚安清华第五期"
person_in_charge = "张智帅"

# ************************************************************************
# 在这里修改文件名
# ************************************************************************


# 命名格式：小伙伴计划-2020春晚安清华第五期获奖名单-张智帅-20101021.xlsx
target_filename = f"小伙伴计划-{activity_name}获奖名单-{person_in_charge}-{date_str}.xlsx"

# Step1：使用pandas处理源数据，取出有效名单
data = pd.read_excel("打卡统计.xls")
winners = data[data["打卡天数"] == 21]
namelist = winners.loc[:, ["圈子昵称", "微信昵称", "打卡天数"]]
namelist.to_excel(target_filename, index=None)

logger.info("Generated the xlsx file successfully!")

# Step2：使用xlwings调格式
app = xw.App(visible=False, add_book=False)
wb = app.books.open(target_filename)
sht = wb.sheets["Sheet1"]
rng = sht.used_range

# ************************************************************************
# 在这里修改格式
# ************************************************************************

# 修改字体、大小、排列
rng.api.Font.Name = "宋体"
rng.api.Font.Size = 14
rng.api.HorizontalAlignment = -4108

# 套用表格格式
tbl = sht.api.ListObjects.add()  # Adds table to Excel (returning a Table)
tbl.TableStyle = "TableStyleMedium4"  # Set table styling

# ************************************************************************
# 在这里修改格式
# ************************************************************************


sht.autofit()

# 保存并退出xlwings
wb.save()
wb.close()
app.quit()

logger.info("Formatted the xlsx file successfully!")
