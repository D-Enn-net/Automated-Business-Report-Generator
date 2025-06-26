# main.py
import time
import shutil
import os
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Import custom modules
from config import INPUT_DATA_PATH, FILE_PREFIX_TO_WATCH, ARCHIVE_PATH
from data_processing import analyze_sales_data
from visualization import create_sales_visualization
from report_builder import create_word_report, create_pdf_report

def process_new_sales_file(file_path):
    """The main workflow to run when a new file is detected."""
    print("="*50)
    print(f"New sales file detected: {file_path}")
    print("Starting automated report generation process...")
    
    metrics = analyze_sales_data(file_path)
    
    if metrics:
        # Generate Reports
        chart_file_path = create_sales_visualization(metrics)
        create_word_report(metrics, chart_file_path)
        create_pdf_report(metrics, chart_file_path)
        print("\nAll reports generated successfully.")

        # Archive the Processed File
        print(f"Archiving processed file: {os.path.basename(file_path)}")
        try:
            destination_path = os.path.join(ARCHIVE_PATH, os.path.basename(file_path))
            shutil.move(file_path, destination_path)
            print("File archived successfully.")
        except Exception as e:
            print(f"Error: Could not archive file. {e}")
            
        print("\nFull process completed successfully for the new file.")
    else:
        print("\nCould not process the file.")
    print("="*50)

class MyEventHandler(FileSystemEventHandler):
    """Custom event handler that reacts to new file creation."""
    def on_created(self, event):
        if not event.is_directory and FILE_PREFIX_TO_WATCH in event.src_path:
            process_new_sales_file(event.src_path)

if __name__ == "__main__":
    event_handler = MyEventHandler()
    observer = Observer()
    observer.schedule(event_handler, path=INPUT_DATA_PATH, recursive=True)
    
    print(f"Starting file system watcher in '{INPUT_DATA_PATH}' for files starting with '{FILE_PREFIX_TO_WATCH}'...")
    print("Press CTRL+C to stop.")
    
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()