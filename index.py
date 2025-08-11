import cv2
import pytesseract
import numpy as np
import csv
import os

# Load image
image = cv2.imread('table.png')

# Convert to grayscale
gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

# Binary inverse threshold (text = white)
_, thresh = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY_INV)

# Detect horizontal lines
h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
h_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, h_kernel)

# Detect vertical lines
v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40))
v_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, v_kernel)

# Combine lines to get table grid
grid = cv2.add(h_lines, v_lines)

# Find contours of cells
contours, _ = cv2.findContours(grid, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

# Sort contours top-to-bottom, then left-to-right
def sort_contours(cnts):
    cnts = sorted(cnts, key=lambda c: cv2.boundingRect(c)[1])  # sort by y
    rows = []
    current_row = []
    last_y = None
    for c in cnts:
        x, y, w, h = cv2.boundingRect(c)
        if last_y is None or abs(y - last_y) < 10:
            current_row.append(c)
            last_y = y
        else:
            rows.append(sorted(current_row, key=lambda c: cv2.boundingRect(c)[0]))
            current_row = [c]
            last_y = y
    rows.append(sorted(current_row, key=lambda c: cv2.boundingRect(c)[0]))
    return rows

rows = sort_contours(contours)

# OCR each cell
table_data = []
for row in rows:
    row_data = []
    for cell in row:
        x, y, w, h = cv2.boundingRect(cell)
        if w < 20 or h < 15:  # skip small boxes
            continue
        roi = image[y:y+h, x:x+w]
        roi_gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
        _, roi_thresh = cv2.threshold(roi_gray, 180, 255, cv2.THRESH_BINARY_INV)
        text = pytesseract.image_to_string(
            roi_thresh,
            config='--psm 6 -c tessedit_char_whitelist=0123456789.:+-ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'
        ).strip()
        row_data.append(text)
    if row_data:
        table_data.append(row_data)

# Save to CSV
output_file = os.path.join(os.getcwd(), "stealth_triggers.csv")
with open(output_file, "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    for row in table_data:
        writer.writerow(row)

print(f"✅ Extracted {len(table_data)} rows saved to {output_file}")
