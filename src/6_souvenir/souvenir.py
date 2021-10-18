# coding: utf8
import os
from io import BytesIO

import pandas as pd
import requests
from PIL import Image


def download_img(img_url, name):
    r = requests.get(img_url, stream=True)
    if r.status_code == 200:
        image = Image.open(BytesIO(r.content))

        if image.mode == "RGBA":
            name = name[:-4] + ".png"

        image.save(name, quality=95, subsampling=0)

    del r
    return image


def get_imgs(excel_name, sheet_name="打卡活动", RESULTS_DIR="results"):

    data = pd.read_excel(excel_name, sheet_name, header=0).T

    for i in range(200):
        usrname = data[i]["用户昵称"]
        urls = data[i]["图片链接"].split(", ")
        print(i, usrname)

        for pic_index in range(len(urls) - 1):
            if pic_index >= (len(urls) - 1) / 2:
                continue

            url = urls[pic_index]

            if url[-7:] == "unknown":
                continue

            image_type = url[-4:]  # 图片类型，可能是.jpg也有可能是.png

            name = usrname + "_" + str(pic_index + 1) + image_type
            target = os.path.join(RESULTS_DIR, name)
            download_img(url, target)


if __name__ == "__main__":

    RESULTS_DIR = "results"
    os.path.exists(RESULTS_DIR) or os.makedirs(RESULTS_DIR)

    excel_name = "【学协】晚安清华·第三期圈子打卡日记.xls"
    get_imgs(excel_name)
