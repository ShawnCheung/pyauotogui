#!/bin/bash
cd /home/shawn/disk/MyPassport/pyautogui/
source activate yolov5
cal >> ./logs/dianming_zrx.txt
date >> ./logs/dianming_zrx.txt
/home/shawn/.pyenv/versions/yolov5/bin/python dianming.py --index 1 >> /home/shawn/disk/MyPassport/pyautogui/logs/dianming.txt
