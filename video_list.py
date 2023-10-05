import os
from moviepy.editor import VideoFileClip
from openpyxl import Workbook
from tqdm import tqdm  # Importa tqdm per la barra di avanzamento
import gc  # Importa il modulo gc per la gestione della memoria












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
            "resolution": video.size[0] if video.size else None,
            "errore": None  # Inizialmente nessun errore
        }
    except Exception as e:
        print(f"Errore nel leggere il file {file_path}: {e}")
        return {
            "name": os.path.basename(file_path),
            "errore": str(e),
            "sizeByte": 0
        }

def explore_directory(directory, batch_size=500):
    video_info_list = []
    total_files = 0
    processed_files = 0

    for root, dirs, files in os.walk(directory):
        total_files += len(files)

    progress_bar = tqdm(total=total_files, desc="Progresso", unit="file")

    for root, dirs, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            if file_path.lower().endswith(('.mp4', '.avi', '.mkv', '.mov')):
                video_info = get_video_info(file_path)
                video_info_list.append(video_info)
                processed_files += 1
                progress_bar.update(1)

                # Libera la memoria ogni 500 file
                if processed_files % batch_size == 0:
                    print("Liberando la memoria...")
                    gc.collect()

    progress_bar.close()
    return video_info_list

def create_excel(video_info_list, excel_path):
    wb = Workbook()
    ws = wb.active

    headers = ["name", "errore", "size", "resolution"]
    ws.append(headers)

    sorted_video_info_list = sorted(video_info_list, key=lambda x: float(x["sizeByte"]), reverse=True)

    for video_info in sorted_video_info_list:
        row = [video_info.get(header, None) for header in headers]
        ws.append(row)

    wb.save(excel_path)

if __name__ == "__main__":
    directory_path = input("Inserisci il percorso della cartella: ")
    excel_file_path = "output.xlsx"

    video_info_list = explore_directory(directory_path)
    
    if video_info_list:
        create_excel(video_info_list, excel_file_path)
        print(f"File Excel creato con successo: {excel_file_path}")
    else:
        print("Nessun file video trovato nella cartella specificata.")
