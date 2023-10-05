import os
from moviepy.editor import VideoFileClip
from openpyxl import Workbook

"""
Create a xlxs file with all video file inside the directory_path and his subfolder,
sorted by syze

Before run ensure dependency are installed

pip install openpyxl moviepy imageio

"""

def convert_size(size_bytes):
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.2f} {unit}"
        size_bytes /= 1024.0

def get_video_info(file_path):
    try:
        video = VideoFileClip(file_path)
        return {
            "size": convert_size(os.path.getsize(file_path)),
            "sizeByte": os.path.getsize(file_path),
            "name": os.path.basename(file_path),
            "resolution": str(video.size[0]) + 'x' + str(video.size[1]) if video.size else None
        }
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")
        return None

def explore_directory(directory, batch_size=500):
    video_info_list = []
    total_files = 0
    processed_files = 0

    print("Collecting data...")
    for root, dirs, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            if file_path.lower().endswith(('.mp4', '.avi', '.mkv', '.mov')):
                total_files += 1

    for root, dirs, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            if file_path.lower().endswith(('.mp4', '.avi', '.mkv', '.mov')):
                print(f"Analizzando il file: {file_path} - {processed_files + 1}/{total_files} ({(processed_files + 1) / total_files * 100:.2f}%)")
                video_info = get_video_info(file_path)
                if video_info:
                    video_info_list.append(video_info)
                    processed_files += 1

                    # Libera la memoria ogni 500 file
                    if processed_files % batch_size == 0:
                        print("Liberando la memoria...")
                        gc.collect()

    return video_info_list

def create_excel(video_info_list, excel_path):
    wb = Workbook()
    ws = wb.active

    headers = ["size", "name", "resolution"]
    ws.append(headers)
    
    sorted_video_info_list = sorted(video_info_list, key=lambda x: float(x["sizeByte"]), reverse=True)

    for video_info in sorted_video_info_list:
        row = [video_info[header] for header in headers]
        ws.append(row)

    wb.save(excel_path)

if __name__ == "__main__":
    directory_path = input("Write folder path: ")
    excel_file_path = "output.xlsx"

    video_info_list = explore_directory(directory_path)
    create_excel(video_info_list, excel_file_path)