import time
from PPTMemory import PPTMemoryEditor

editor = PPTMemoryEditor()
while not editor.attach():
    print("Waiting for PPT...")
    time.sleep(1)

while True:
    editor.set_next(["I"] * 5)
    time.sleep(0.2)
