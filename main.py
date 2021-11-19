# xlrd 1.2.0
import xlrd
import os
from pathlib import Path

file_location='database.xlsx'

workbook = xlrd.open_workbook(file_location)
sheet_video = workbook.sheet_by_name('video-test')
video_titles = sheet_video.col_values(0, start_rowx=1)
video_ids = sheet_video.col_values(1, start_rowx=1)
video_ids_str = [str(int(id)) for id in video_ids]

# Task 1.2: Match IDs with Titles
id_dic = {}
# print(len(video_ids))
for i in range (len(video_ids)):
    id_dic[video_ids_str[i]] = video_titles[i]


# Task 1.3: Rename each file from its ID to its Title
path = 'test'
os.chdir(path)

# print(list_os)
for f in os.listdir():
    id_video, suffer = os.path.splitext(f)
    for ids, titles in id_dic.items():
        if ids == id_video:
            name = f'{titles}{suffer}'
            # print(name)
            os.rename(f, name)

# Task 2: Write code to move the files into the folder of each genre
sheet_main_genres = workbook.sheet_by_name('main-genres')
list_main_genres = sheet_main_genres.col_values(0, start_rowx=1)
# print(list_main_genres)
# tao thu muc
# os.chdir(path)

for diractory in list_main_genres:
    diractory_path = Path(diractory)
    try:
        os.mkdir(diractory_path)
    except:
        pass

for f in os.scandir():
    file_path = Path(f)
    if f.is_dir():
        continue
    f = str(f)
    for genre in list_main_genres:
        if genre in f:
            genrePath = Path(genre)
            file_path.rename(genrePath.joinpath(file_path))
            # print(file_path.joinpath(genrePath))
