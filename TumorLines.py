#TUMOR DETECTION PROGRAM (PYTHON)
#DESIGNED BY DREW LAWLESS
#BELLAMKONDA/MOKARRAM LAB (EMORY UNIVERSITY)
#created: 01/22/2024 | last updated: 03/04/2024

import os
import openpyxl
import numpy as np
import cv2 as cv
import matplotlib
from matplotlib import pyplot as plt
import shutil
import openpyxl

main_path = '/Users/drew/Desktop/Research/Image Analysis'                                                                                 
input_path = '/Users/drew/Desktop/Research/Image Analysis/Input'
output_path = '/Users/drew/Desktop/Research/Image Analysis/Output'
purple_path = '/Users/drew/Desktop/Research/Image Analysis/Purple'
excel_path = '/Users/drew/Desktop/Research/Image Analysis/Python Image Analysis.xlsx'
processed_path = '/Users/drew/Desktop/Research/Image Analysis/Processed'

wb = openpyxl.load_workbook(excel_path)
sheet1 = wb['February in vivo']

imageCounter = 1
row_counter = 2
row_counter2 = 2
for filename in os.listdir(input_path):
    if '.tif' in filename:
        voxel_ratio = 0

        #begin image and open image
        print("ANAlYZING IMAGE #" + str(imageCounter) + "...")
        file = os.path.join(input_path, filename)
        oG = cv.imread(file)
        image = oG

        #set ranges for circle colors
        lower_white = np.array([200, 200, 200], dtype=np.uint8)
        upper_white = np.array([255, 255, 255], dtype=np.uint8)

        lower_green = np.array([0, 100, 0], dtype=np.uint8)
        upper_green = np.array([45, 255, 45], dtype=np.uint8)

        lower_yellow = np.array([0, 100, 100], dtype=np.uint8)
        upper_yellow = np.array([50, 255, 255], dtype=np.uint8)

        lower_blue = np.array([0, 0, 50], dtype=np.uint8)
        upper_blue = np.array([100, 100, 255], dtype=np.uint8)

        lower_red = np.array([50, 0, 0], dtype=np.uint8)
        upper_red = np.array([255, 100, 100], dtype=np.uint8)

        lower_purple = np.array([100, 0, 100], dtype=np.uint8)
        upper_purple = np.array([255, 50, 255], dtype=np.uint8)

        #set masks for finding contours
        white_mask = cv.inRange(image, lower_white, upper_white)
        green_mask = cv.inRange(image, lower_green, upper_green)
        yellow_mask = cv.inRange(image, lower_yellow, upper_yellow)
        blue_mask = cv.inRange(image, lower_blue, upper_blue)
        red_mask = cv.inRange(image, lower_red, upper_red)
        purple_mask = cv.inRange(image, lower_purple, upper_purple)

        #contours for each color
        white_contours, hierarchy = cv.findContours(white_mask, cv.RETR_EXTERNAL, cv.CHAIN_APPROX_SIMPLE)
        green_contours, _ = cv.findContours(green_mask, cv.RETR_EXTERNAL, cv.CHAIN_APPROX_SIMPLE)
        yellow_contours, hierarchy = cv.findContours(yellow_mask, cv.RETR_TREE, cv.CHAIN_APPROX_SIMPLE)
        blue_contours, _ = cv.findContours(blue_mask, cv.RETR_EXTERNAL, cv.CHAIN_APPROX_SIMPLE)
        red_contours, _ = cv.findContours(red_mask, cv.RETR_EXTERNAL, cv.CHAIN_APPROX_SIMPLE)
        purple_contours, _ = cv.findContours(purple_mask, cv.RETR_EXTERNAL, cv.CHAIN_APPROX_SIMPLE)

        #check for white dots
        white_contours2 = []
        if len(white_contours) >= 3:
            for contour in white_contours:
                x, y, w, h = cv.boundingRect(contour)
                if 35 <= w <= 85:
                    cv.line(image, (x, y), (x + w, y), (255, 255, 255), 2)
                    voxel_ratio = 500 / w
                    print("Voxel Ratio: " + str(voxel_ratio) + " µm/px.")
                else:
                    white_contours2.append(contour)
                    
            sorted_contours = sorted(white_contours2, key=cv.contourArea, reverse=True)[:2]
            square_centers = [tuple(np.mean(contour, axis=0, dtype=int)[0]) for contour in sorted_contours]

            for center in square_centers:
                cv.circle(image, center, 5, (255, 255, 255), -1)

            cv.line(image, square_centers[0], square_centers[1], (255, 255, 255), 2)

            square1_center = np.array(square_centers[0])
            square2_center = np.array(square_centers[1])
            extension_length = 100
            line_vector = square2_center - square1_center
            endpoint1 = square1_center - extension_length * line_vector / np.linalg.norm(line_vector)
            endpoint2 = square2_center + extension_length * line_vector / np.linalg.norm(line_vector)
            cv.line(image, (int(endpoint1[0]), int(endpoint1[1])), (int(endpoint2[0]), int(endpoint2[1])), (255, 255, 255), 2)

            #checks for green circle
            if len(green_contours) > 0:
                largest_green_contour = max(green_contours, key=cv.contourArea)
                green_center = tuple(np.mean(largest_green_contour, axis=0, dtype=int)[0])
                cv.circle(image, green_center, 5, (255, 255, 255), -1)

                vector = np.array(square_centers[1]) - np.array(square_centers[0])
                perpendicular_vector = np.array([-vector[1], vector[0]])
                endpoint1 = (int(green_center[0] - 4 * perpendicular_vector[0]), int(green_center[1] - 4 * perpendicular_vector[1]))
                endpoint2 = (int(green_center[0] + 4 * perpendicular_vector[0]), int(green_center[1] + 4 * perpendicular_vector[1]))
                cv.line(image, green_center, endpoint1, (255, 255, 255), 2)
                cv.line(image, green_center, endpoint2, (255, 255, 255), 2)
                print("Reference cross drawn.")

            else:
                print("No green circles detected.")

        else:
            print("No white contours found.")

        #draw extra lines
        if len(green_contours) > 0 and len(yellow_contours) > 0 and len(blue_contours) > 0 and len(red_contours) > 0:
                
            #Check for all contours and add reference center circle
            largest_yellow_contour = max(yellow_contours, key=cv.contourArea)
            largest_green_contour = max(green_contours, key=cv.contourArea)
            largest_blue_contour = max(blue_contours, key=cv.contourArea)
            largest_red_contour = max(red_contours, key=cv.contourArea)

            yellow_center, yellow_radius = cv.minEnclosingCircle(largest_yellow_contour)[:2]
            green_center, green_radius = cv.minEnclosingCircle(largest_green_contour)[:2]
            blue_center, blue_radius = cv.minEnclosingCircle(largest_blue_contour)[:2]
            red_center, red_radius = cv.minEnclosingCircle(largest_red_contour)[:2]

            yellow_center = (int(yellow_center[0]), int(yellow_center[1]))
            green_center = (int(green_center[0]), int(green_center[1]))
            blue_center = (int(blue_center[0]), int(blue_center[1]))
            red_center = (int(red_center[0]), int(red_center[1]))

            cv.circle(image, yellow_center, 4, (255, 255, 255), -1)
            cv.circle(image, green_center, 4, (255, 255, 255), -1)
            cv.circle(image, blue_center, 4, (255, 255, 255), -1)
            cv.circle(image, red_center, 4, (255, 255, 255), -1)

            #check for purple contours and draw each images' lines
            if len(purple_contours) > 0:
                os.chdir(purple_path)
                purple_counter = 1
                for contour in purple_contours:
                    contour_area = cv.contourArea(contour)
                    if 5 <= cv.contourArea(contour) <= 3000:
                        print("Purple Contour " + str(purple_counter) + " added.")
                        image_copy = image.copy()
                        purple_center, purple_radius = cv.minEnclosingCircle(contour)[:2]
                        purple_center = (int(purple_center[0]), int(purple_center[1]))
                        cv.circle(image, purple_center, 2, (255, 255, 255), -1)
                        cv.line(image_copy, purple_center, blue_center, (255, 255, 255), 2)
                        cv.line(image_copy, purple_center, red_center, (255, 255, 255), 2)
                        cv.line(image_copy, purple_center, green_center, (255, 255, 255), 2)
                        cv.imwrite("p" + str(purple_counter) + "_" + str(filename), image_copy)

                        contour_radius = np.sqrt(contour_area / 3.1415926)
                        purple_red_length = np.sqrt((red_center[0] - purple_center[0])**2 + (red_center[1] - purple_center[1])**2)
                        purple_blue_length = np.sqrt((blue_center[0] - purple_center[0])**2 + (blue_center[1] - purple_center[1])**2)
                        purple_green_length = np.sqrt((green_center[0] - purple_center[0])**2 + (green_center[1] - purple_center[1])**2)

                        sheet1['AJ' + str(row_counter2)] = str(filename)
                        sheet1['AK' + str(row_counter2)] = contour_area * voxel_ratio
                        sheet1['AL' + str(row_counter2)] = purple_radius * voxel_ratio
                        sheet1['AM' + str(row_counter2)] = purple_red_length * voxel_ratio
                        sheet1['AN' + str(row_counter2)] = purple_blue_length * voxel_ratio
                        sheet1['AO' + str(row_counter2)] = purple_green_length * voxel_ratio
                        
                        sheet1['AQ' + str(row_counter2)] = str(filename)
                        sheet1['AR' + str(row_counter2)] = contour_area
                        sheet1['AS' + str(row_counter2)] = purple_radius
                        sheet1['AT' + str(row_counter2)] = purple_red_length
                        sheet1['AU' + str(row_counter2)] = purple_blue_length
                        sheet1['AV' + str(row_counter2)] = purple_green_length

                        row_counter2 += 1                      
                        purple_counter += 1
            else:
                print("No purple contours detected.")

            os.chdir(main_path)
   
            #draw all other lines
            cv.line(image, green_center, yellow_center, (255, 255, 255), 2)
            cv.line(image, red_center, blue_center, (255, 255, 255), 2)
            cv.line(image, green_center, blue_center, (255, 255, 255), 2)
            cv.line(image, green_center, red_center, (255, 255, 255), 2)
            cv.line(image, yellow_center, blue_center, (255, 255, 255), 2)
            cv.line(image, yellow_center, red_center, (255, 255, 255), 2)

            #calculates tumor distance from reference line
            green_to_yellow_vector = np.array(square_centers[1]) - np.array(square_centers[0])
            perpendicular_vector = np.array([-green_to_yellow_vector[1], green_to_yellow_vector[0]])
            endpoint1 = (int(green_center[0] - 4 * perpendicular_vector[0]), int(green_center[1] - 4 * perpendicular_vector[1]))
            endpoint2 = (int(green_center[0] + 4 * perpendicular_vector[0]), int(green_center[1] + 4 * perpendicular_vector[1]))
            distance = np.abs(np.cross(green_to_yellow_vector, np.array(green_center) - np.array(endpoint1))) / np.linalg.norm(green_to_yellow_vector)

            #creates line lengths and adds to Excel sheet
            red_edge_length = np.sqrt((red_center[0] - yellow_center[0])**2 + (red_center[1] - yellow_center[1])**2) - yellow_radius
            blue_edge_length = np.sqrt((blue_center[0] - yellow_center[0])**2 + (blue_center[1] - yellow_center[1])**2) - yellow_radius
            red_green_length = np.sqrt((red_center[0] - green_center[0])**2 + (red_center[1] - green_center[1])**2)
            blue_green_length = np.sqrt((blue_center[0] - green_center[0])**2 + (blue_center[1] - green_center[1])**2)
            red_blue_length = np.sqrt((red_center[0] - blue_center[0])**2 + (red_center[1] - blue_center[1])**2)
            green_yellow_length = np.sqrt((yellow_center[0] - green_center[0])**2 + (yellow_center[1] - green_center[1])**2)

            sheet1['A' + str(row_counter)] = str(filename)
            sheet1['B' + str(row_counter)] = cv.contourArea(largest_green_contour) * voxel_ratio
            sheet1['C' + str(row_counter)] = green_radius * voxel_ratio
            sheet1['D' + str(row_counter)] = cv.contourArea(largest_yellow_contour) * voxel_ratio
            sheet1['E' + str(row_counter)] = yellow_radius * voxel_ratio
            sheet1['F' + str(row_counter)] = cv.contourArea(largest_red_contour) * voxel_ratio
            sheet1['G' + str(row_counter)] = red_radius * voxel_ratio
            sheet1['H' + str(row_counter)] = cv.contourArea(largest_blue_contour) * voxel_ratio
            sheet1['I' + str(row_counter)] = blue_radius * voxel_ratio
            sheet1['J' + str(row_counter)] = red_edge_length * voxel_ratio
            sheet1['K' + str(row_counter)] = blue_edge_length * voxel_ratio
            sheet1['L' + str(row_counter)] = red_green_length * voxel_ratio
            sheet1['M' + str(row_counter)] = blue_green_length * voxel_ratio
            sheet1['N' + str(row_counter)] = red_blue_length * voxel_ratio
            sheet1['O' + str(row_counter)] = green_yellow_length * voxel_ratio
            sheet1['P' + str(row_counter)] = distance * voxel_ratio
            sheet1['Q' + str(row_counter)] = voxel_ratio

            sheet1['S' + str(row_counter)] = str(filename)
            sheet1['T' + str(row_counter)] = cv.contourArea(largest_green_contour)
            sheet1['U' + str(row_counter)] = green_radius
            sheet1['V' + str(row_counter)] = cv.contourArea(largest_yellow_contour)
            sheet1['W' + str(row_counter)] = yellow_radius
            sheet1['X' + str(row_counter)] = cv.contourArea(largest_red_contour)
            sheet1['Y' + str(row_counter)] = red_radius
            sheet1['Z' + str(row_counter)] = cv.contourArea(largest_blue_contour)
            sheet1['AA' + str(row_counter)] = blue_radius
            sheet1['AB' + str(row_counter)] = red_edge_length
            sheet1['AC' + str(row_counter)] = blue_edge_length
            sheet1['AD' + str(row_counter)] = red_green_length
            sheet1['AE' + str(row_counter)] = blue_green_length
            sheet1['AF' + str(row_counter)] = red_blue_length
            sheet1['AG' + str(row_counter)] = green_yellow_length
            sheet1['AH' + str(row_counter)] = distance
            row_counter += 1
            wb.save(excel_path)

            print("Image #" + str(imageCounter) + " finished, all lines drawn.")

        else:
            print("Not enough circles detected, No extra lines drawn!")

        print("\n-----------------\n")

        #move previous image and start new image
        shutil.move(file, os.path.join(processed_path, filename))
        os.chdir(output_path)
        cv.imwrite(filename, image)
        os.chdir(main_path)
        imageCounter += 1

print("IMAGE ANALYSIS COMPLETE!")