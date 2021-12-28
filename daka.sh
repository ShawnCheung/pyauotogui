#!/bin/bash
cd /home/shawn/disk/MyPassport/pyautogui/
source activate yolov5
cal >> ./logs/daka.txt
date >> ./logs/daka.txt
/home/shawn/.pyenv/versions/yolov5/bin/python daka.py --index 1 >> /home/shawn/disk/MyPassport/pyautogui/logs/daka.txt
