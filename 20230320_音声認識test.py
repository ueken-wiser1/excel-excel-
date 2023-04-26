import speech_recognition as sr
import os
from openpyxl import Workbook
from pydub import AudioSegment

# 音声認識オブジェクトを作成
recognizer = sr.Recognizer()

# 音声ファイルのリスト
def list_audio_files(folder_path, audio_extension=".m4a"):
    audio_files = []
    
    for file in os.listdir(folder_path):
        if file.endswith(audio_extension):
            audio_files.append(os.path.join(folder_path, file))

    return audio_files


folder_path = "C:/Users/touko/Documents\サウンド レコーディング"
audio_files = list_audio_files(folder_path)
print(audio_files)

def convert_m4a_to_wav(file_path):
    audio = AudioSegment.from_file(file_path, format="m4a")
    wav_file_path = os.path.splitext(file_path)[0] + ".wav"
    audio.export(wav_file_path, format="wav")
    return wav_file_path

def recognize_audio_file(file_path):
    wav_file_path = convert_m4a_to_wav(file_path)
    with sr.AudioFile(wav_file_path) as source:
        audio_data = recognizer.record(source)

        try:
            text = recognizer.recognize_google(audio_data, language="ja-JP")
            print(f"{file_path} の音声認識結果: {text}")
            return text
        except sr.UnknownValueError:
            print(f"{file_path} の音声認識できませんでした")
        except sr.RequestError as e:
            print(f"{file_path} の音声認識サービスへのリクエストに失敗しました; {e}")

# Excelファイルの作成と初期設定
wb = Workbook()
ws = wb.active

# 音声認識結果を書き込む列
column = 1

# 各音声ファイルに対して音声認識を実行
for index, audio_file in enumerate(audio_files, start=1):
    if os.path.isfile(audio_file):
        text = recognize_audio_file(audio_file)
        if text:
            # 音声認識結果をExcelファイルの所定のセルに書き込む
            ws.cell(row=index, column=column, value=text)
            
            # 音声ファイルを削除
            os.remove(audio_file)
    else:
        print(f"{audio_file} が見つかりませんでした。")

# Excelファイルを保存
wb.save("recognized_text.xlsx")
