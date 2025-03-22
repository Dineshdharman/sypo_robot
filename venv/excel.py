import openpyxl
from openpyxl.styles import Font
import os

# Define file path (Change this to your shared Google Drive/OneDrive path)
file_path = r"C:\Users\HP\Desktop\SymposiumRobot\venv\symposium_event.xlsx"  # Modify this based on your shared location

# Check if the file exists
if os.path.exists(file_path):
    # Load existing file for real-time update
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
else:
    # Create a new workbook and worksheet
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Events"

    # Add headers
    headers = ["Event Name", "Type", "Floor", "Room No"]
    ws.append(headers)
    
    # Make headers bold
    for col in range(1, len(headers) + 1):
        ws.cell(row=1, column=col).font = Font(bold=True)

# Event data
event_data = [
    ["Debugging", "Technical", 2, 205],
    ["Paper Presentation", "Technical", 3, 307],
    ["Coding Challenge", "Technical", 4, 410],
    ["Photography", "Non-Technical", 1, 101],
    ["Quiz", "Non-Technical", 1, 102]
]

# Append data if not already present
existing_events = [ws.cell(row=i, column=1).value for i in range(2, ws.max_row + 1)]
for event in event_data:
    if event[0] not in existing_events:  # Avoid duplicate entries
        ws.append(event)

# Save the file
wb.save(file_path)
print(f"Excel file updated at {file_path}")

# Instructions for sharing
print("Share this Excel file via Google Drive or OneDrive for real-time access.")
