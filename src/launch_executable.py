import subprocess
import sys
import time

if __name__ == "__main__":
    executable_path = str(sys.argv[1])
    
    # Give time for any programs to close
    time.sleep(2)
    subprocess.run(executable_path)