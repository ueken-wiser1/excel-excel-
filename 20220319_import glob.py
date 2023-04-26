import glob

files = glob.glob("C:/Users/touko/OneDrive/自動化用/02.ラーメン画像/ラーメン画像/*.JPG")
for file in files:
    print(file)