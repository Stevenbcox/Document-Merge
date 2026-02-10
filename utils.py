# Utility functions for the application
def progress_callback(progress_queue, value):
    if progress_queue is None:
        return  # no queue, nothing to do
    progress_queue.put(value)  # value between -1 and 1
