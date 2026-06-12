import shutil
import os

def get_storage_info():
    total, used, free = shutil.disk_usage('/')
    return total, used, free

def format_size(size):
    for unit in ['', 'K', 'M', 'G', 'T', 'P', 'E', 'Z']:
        if abs(size) < 1024.0:
            return f"{size:3.1f}{unit}B"
        size /= 1024.0
    return f"{size:.1f}YB"

def print_storage_info():
    total, used, free = get_storage_info()
    print(f"Total Storage: {format_size(total)}")
    print(f"Used Storage: {format_size(used)}")
    print(f"Free Storage: {format_size(free)}")

def main():
    print_storage_info()

if __name__ == "__main__":
    main()