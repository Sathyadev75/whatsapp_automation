import streamlit as st
import time
import pyautogui
import pygetwindow as gw
import pandas as pd
import matplotlib.pyplot as plt
import random

# Function to check if WhatsApp window is active
def is_whatsapp_window_active():
    active_window = gw.getActiveWindow()
    return active_window is not None and "WhatsApp" in active_window.title

# Function to convert Excel to image
def convert_excel_to_image():
    df = pd.read_excel('output_data.xlsx', engine='openpyxl')

    fig, ax = plt.subplots(figsize=(6, 4), dpi=100)
    ax.axis('off')

    table = ax.table(cellText=df.values, colLabels=df.columns, cellLoc='center', loc='center')

    for (i, j), cell in table._cells.items():
        if i == 0:
            cell.set_height(0.2)
            cell.set_fontsize(60)
            cell.set_width(0.12)

            cell_text = cell.get_text().get_text()
            if ' ' in cell_text:
                cell_text = cell_text.replace(' ', '\n')
                cell.get_text().set_text(cell_text)
        else:
            cell.set_width(0.12)
            cell.set_fontsize(60)
            cell.set_text_props(weight='bold', color='black')

        if j == 0 or j == 1:
            cell.set_width(0.17)
        if i != 0:
            cell.set_fontsize(80)

        if i == len(df):
            cell.set_facecolor('#87CEEB')
            cell.set_text_props(weight='bold', color='black')

    for key, cell in table._cells.items():
        if key[0] == 0:
            cell.set_text_props(weight='bold', color='w')
            cell.set_facecolor('#1f77b4')
            cell.set_edgecolor('k')
        else:
            cell.set_edgecolor('k')

    table.auto_set_font_size(False)
    table.scale(15, 17)

    image_path = f'table_image{round(random.random(), 2)}.png'
    plt.savefig(image_path, bbox_inches='tight', pad_inches=0)
    st.success("Finished image creation")
    return image_path

# Streamlit app
def main():
    st.title("WhatsApp Automation")

    # WhatsApp Number Input
    whatsapp_number = st.text_input("Enter WhatsApp Number:")

    # File Upload
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls", "csv"])

    if st.button("Perform WhatsApp Actions"):
        if whatsapp_number and uploaded_file:
            # Perform WhatsApp actions using the provided functions
            try:
                image_path = convert_excel_to_image()
                st.success(f"Executed for {whatsapp_number} and created image: {image_path}")
            except Exception as e:
                st.error(f"Error: {str(e)}")
        else:
            st.warning("Please enter a WhatsApp number and upload a file.")

if __name__ == '__main__':
    main()
