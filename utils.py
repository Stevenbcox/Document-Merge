# Utility functions for the application

def progress_callback(progress_queue, value):
    progress_queue.put(value)  # value between -1 and 1
    pass
