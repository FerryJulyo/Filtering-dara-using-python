import pandas as pd
import threading
import itertools
import sys
import time

def show_spinner(message, stop_event):
    spinner = itertools.cycle(['|', '/', '-', '\\'])
    while not stop_event.is_set():
        sys.stdout.write(f'\r{message} {next(spinner)}')
        sys.stdout.flush()
        time.sleep(0.1)
    sys.stdout.write(f'\r{message} ✓\n')

def process_file(file_path, output_path):
    stop_event = threading.Event()
    message = f"Processing {file_path}"
    spinner_thread = threading.Thread(target=show_spinner, args=(message, stop_event))
    spinner_thread.start()

    excel_file = pd.ExcelFile(file_path)
    df_list = []
    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        df['SheetName'] = sheet_name
        df_list.append(df)
    combined_df = pd.concat(df_list, ignore_index=True)
    combined_df.to_excel(output_path, index=False)

    stop_event.set()
    spinner_thread.join()

# Proses file pertama
process_file('Master.xlsx', 'Combined/MasterCombined.xlsx')

# Proses file kedua
process_file('TopHit.xlsx', 'Combined/TopHitCombined.xlsx')
