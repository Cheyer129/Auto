from pathlib import Path
import os 


CurrentDirectory = Path.cwd()

users = os.listdir('Toscano Files/USERS')

print(users)