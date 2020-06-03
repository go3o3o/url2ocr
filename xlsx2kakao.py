import cv2
import json
import requests
import sys
import os
import re
import hashlib
import configparser
import shutil
import logging

import numpy as np

from openpyxl import load_workbook
from PIL import Image
from urllib.request import urlopen
from urllib import parse

# Config Parser 초기화
config = configparser.ConfigParser()
# Config File 읽기
config.read(os.path.dirname(os.path.realpath(__file__)) + os.sep + 'envs' + os.sep + 'property.ini')

LIMIT_PX = 1024
LIMIT_BYTE = 1024*1024  # 1MB
LIMIT_BOX = 40

def url_to_image(url, readFlog=cv2.IMREAD_COLOR):
    p = re.compile('[가-힣]+')
    hangeuls = p.findall(url)

    for hangeul in hangeuls:
        url = url.replace(hangeul, parse.quote(hangeul))

    resp = urlopen(url)
    image = np.asarray(bytearray(resp.read()), dtype=np.uint8)
    image = cv2.imdecode(image, readFlog)

    return image

def kakao_ocr_resize(image: str):
    height, width, _ = image.shape

    if LIMIT_PX < height or LIMIT_PX < width:
        ratio = float(LIMIT_PX) / max(height, width)
        image = cv2.resize(image, None, fx=ratio, fy=ratio)
        height, width, _ = height, width, _ = image.shape

        # api 사용전에 이미지가 resize된 경우, recognize시 resize된 image를 사용해야함.
        return image
    return None

def kakao_ocr_detect(image: str, appkey: str):
    API_URL = 'https://kapi.kakao.com/v1/vision/text/detect'

    headers = {'Authorization': 'KakaoAK {}'.format(appkey)}

    jpeg_image = cv2.imencode(".jpg", image)[1]
    data = jpeg_image.tobytes()

    return requests.post(API_URL, headers=headers, files={"file": data})

def kakao_ocr_recognize(image: str, boxes: list, appkey: str):
    API_URL = 'https://kapi.kakao.com/v1/vision/text/recognize'

    headers = {'Authorization': 'KakaoAK {}'.format(appkey)}

    jpeg_image = cv2.imencode(".jpg", image)[1]
    data = jpeg_image.tobytes()

    return requests.post(API_URL, headers=headers, files={"file": data}, data={"boxes": json.dumps(boxes)})

def getFiles(path: str, files: list):
    listdirs = os.listdir(path)

    for listdir in listdirs:
        subpath = path + '/' + listdir
        # rint(subpath)
        if os.path.isdir(subpath):
            getFiles(subpath, files)
        else:
            files.append(subpath)

def md5Generator(text: str):
    enc = hashlib.md5()
    enc.update(text.encode('utf-8'))
    encText = enc.hexdigest()
    return encText

def main():

    appkeys = ['1', '2', '3']

    # xlsx 파일 경로
    xlsxPath = os.path.dirname(os.path.realpath(__file__)) + config['Path']['XlsxPath']

    # result 파일 저장 경로
    resultPath = os.path.dirname(os.path.realpath(__file__)) + config['Path']['ResultPath']

    file = 1
    file_seq = 1

    appkey_seq = 0
    appkey = appkeys[appkey_seq]

    print(" @@@@@@@@@@@@@@@ API KEY %s " % appkey)

    xlsxFiles = []
    getFiles(xlsxPath, xlsxFiles)
    print(" ### xlsx 파일 총 건수 %d " % len(xlsxFiles))

    row_seq = 0

    for xlsxFile in xlsxFiles:
        print("Step #1. xlsx 파일 읽기 ")
        print(" ### %s " % xlsxFile)
        load_wb = load_workbook(xlsxFile)
        load_ws = load_wb.worksheets[0]

        for row in load_ws.rows:
            row_seq += 1

            if row[6].value == "Y":
                try:
                    print("Step #2. 이미지 URL -> OCR %d" % row_seq)
                    print("Step #2-1. %s " % row[4].value)
                    image = url_to_image(row[4].value)

                    resize_image = kakao_ocr_resize(image)
                    if resize_image is not None:
                        image = resize_image
                        print(" ### 원본 대신 리사이즈된 이미지를 사용합니다.")

                    output = kakao_ocr_detect(image, appkey).json()

                    if 'result' in output:
                        boxes = output["result"]["boxes"]
                    else:
                        break

                    boxes = boxes[:min(len(boxes), LIMIT_BOX)]
                    output = kakao_ocr_recognize(image, boxes, appkey).json()

                    if 'result' in output:
                        ocrResult = ' '.join(output["result"]["recognition_words"])
                    else:
                        ocrResult = ''

                    print("Step #2-2. %s " % ocrResult.strip())

                    row[5].value = ocrResult.strip()

                except KeyError:
                    appkey_seq += 1
                    if len(appkeys) <= appkey_seq:
                        print(" @@@@@@@@@@@@@@@ API KEY Expired. %d " % row_seq)
                        break
                    else:
                        appkey = appkeys[appkey_seq]
                        print(" @@@@@@@@@@@@@@@ API KEY %s " % appkey)
                except:
                    pass

                row[6].value = ""

            else:
                row[6].value = ""

        filename = 'result_' + str(file_seq)
        load_wb.save(resultPath + '/' + filename + '.xlsx')
        print("Step #6. 엑셀 파일 저장: %s " % (resultPath + '/' + filename + '.xlsx'))

        file_seq += 1

        if len(appkeys) < appkey_seq:
            print("Step #7. API KEY Expired. %s " % xlsxFile)
            break

    print("Step #7. 끗!!!!!!!!!!!!!!!!!!")


if __name__ == "__main__":
    main()
