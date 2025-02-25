# IMAGE ROTATION PROGRAM (PYTHON)
# DESIGNED BY DREW LAWLESS
# MOKARRAM/BELLAMKONDA LAB (EMORY UNIVERSITY)
# created: 04/22/2024 | last updated: 04/22/2024

import os
import openpyxl
import numpy as np
import cv2 as cv
import matplotlib
from matplotlib import pyplot as plt
import shutil
import openpyxl

main_path = 'C:/Users/bellamkondalabuser/Desktop/Python Image Analysis'
output_path = 'C:/Users/bellamkondalabuser/Desktop/Python Image Analysis/Output'
processed_path = 'C:/Users/bellamkondalabuser/Desktop/Python Image Analysis/Processed'

def rotate():
	files = os.listdir(input_path)
  
    for file_name in files:
        if file_name.endswith(('.tif' or '.jpg')):
            image_path = os.path.join(input_path, file_name)
            img = Image.open(image_path)
            img.save(os.path.join(processed_path, f"rotated_{angle}_{file_name}")) 

            angle = float(input("Enter the angle to rotate the image (in degrees), or enter 0 to skip: "))
            if angle == 0:
                continue

            rotated_img = img.rotate(angle, expand=True)
            rotated_img.save(os.path.join(output_path, f"rotated_{angle}_{file_name}")) 

rotate()







