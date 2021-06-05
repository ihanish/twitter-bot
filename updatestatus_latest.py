import tweepy
import time
import re
import numpy as np
# import docx
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook

auth = tweepy.OAuthHandler('') #add your OAuthHandler
auth.set_access_token('') #add your access_token

api = tweepy.API(auth)
statusfile = 'tweets.txt'
mediafiles = []
media_ids = []
excel_done = 'tweets_done.xlsx'


def status_endinglines(status, status_list):
    status = status.strip('\n')
    author = status.split('\n\n')
    if len(author) > 1:
        pattern = re.compile(author[1])
        f_string = pattern.search(status_list[len(status_list)-1])
        if f_string:
            return status_list
        else:
            copy = status_list[len(status_list)-1]
            # last element is assigned author name
            status_list[len(status_list)-1] = author[1]
            cut = author[1].replace(copy, '')
            new = status_list[len(status_list)-2].strip(cut)
            status_list[len(status_list)-2] = new
            return status_list
    else:
        return status_list


def get_last_tweet():
    tweet = api.user_timeline(id=api.me().id, count=1)[0]
    return tweet.id


def curate_status(status):
    status_list = []

    status = status.strip('\r\n')
    status = status.split('\r\n\r\n')
    i = 0
    while i < len(status[0]):
        if len(status[0][i:]) > 280:
            if i == 0:
                cut = ' '.join(
                    status[0][i: i+277].split(' ')[0:-1])
                add = cut.strip() + '. . .'
            else:
                cut = ' '.join(
                    status[0][i: i+276].split(' ')[0:-1])
                add = '. . .' + cut.strip() + '. . .'

        else:
            cut = ' '.join(status[0][i: i+280].split(' '))
            add = cut.strip()
            if len(status_list) >= 1:
                add = '. . .' + cut.strip()
        i = i + len(cut)
        status_list.append(add.strip())
    if len(status_list[len(status_list)-1]) + len(status[1]) + len('\n\n') <= 280:
        status_list[len(status_list) -
                    1] = status_list[len(status_list)-1] + '\n\n' + status[1]
    else:
        status_list.append(status[1])
    return status_list


def size_of_file(statusfile):
    with open(statusfile, mode='rb') as file:
        file.seek(0, 2)
        size = file.tell()
    return size


def load_status(statusfile, seekfile='seek.txt'):
    with open(seekfile, mode='rb') as s:
        start = int(s.readline().decode('utf-8').strip('SEEK:'))
        t_num = int(s.readline().decode('utf-8').strip('TWEET_NUM:'))
    start = 0
    status = ''
    with open(statusfile, mode='rb') as f:
        if start < size_of_file(statusfile):
            f.seek(start)
            read = True
            while read:
                line = f.readline().decode('utf-8')
                if line.strip() == '*Start*':
                    read = False
                    break
                elif f.tell() == size_of_file(statusfile):
                    break
            read = True
            while read:
                line = f.readline().decode('utf-8')
                if line.strip() == '*End*':
                    read = False
                    break
                elif f.tell() == size_of_file(statusfile):
                    break
                else:
                    status = status + line
            end = f.tell()
        else:
            print("Please update the file with more tweets")
            exit()
    if status in ('' or '\r\n' or ' '):
        print('Nothing to Tweet')
        exit()
    else:
        status_list = curate_status(status)
    return status_list, end, t_num, status


def media_upload(mediafile):
    with open(mediafile, 'rb') as file:
        media_id = api.media_upload(mediafile, file=file)
    return media_id.media_id_string


def update_file(statusfile):
    with open(statusfile, mode='r', encoding='utf-8') as f:
        read = True
        while read:
            line = f.readline()
            with open('tweets_done.txt', mode='a', encoding='utf-8') as done:
                done.write(line)
            if line.strip() == '*End*':
                read = False
                break
        newfile = f.readlines()
    with open(statusfile, mode='w', encoding='utf-8') as n:
        n.writelines(newfile)


def update_excel(status):
    list_excel = status.split('\r\n\r\n')
    list_excel[1] = list_excel[1].strip('-' or '—')
    list_author_book = list_excel[1].split(',')
    list_author_book[1] = list_author_book[1].strip(' ')
    wb = load_workbook(excel_done)
    ws = wb.active
    df = pd.DataFrame({'Tweet': [list_excel[0]],
                       'Book': [list_author_book[0]],
                       'Author': [list_author_book[1]]})
    for r in dataframe_to_rows(df, index=False, header=False):
        ws.append(r)
    wb.save(excel_done)


def update_status(statuslist, media_ids, end, t_num, status):
    api.update_status(status=statuslist[0], media_ids=media_ids)
    if len(statuslist) > 1:
        for i in range(1, len(statuslist)):
            api.update_status(
                status=statuslist[i], in_reply_to_status_id=get_last_tweet(), media_ids=media_ids)
    update_file(statusfile)
    t_num = t_num+1
    print(statuslist)
    print(f'Status Updated/TWEET_NUM:{t_num}')
    with open('seek.txt', mode='w') as file:
        file.write(str(f'SEEK:{end}\nTWEET_NUM:{t_num}'))
    update_excel(status)


def generate_media_ids(mediafiles):
    if len(mediafiles) > 4:
        print('Number of images should be less than 4')
    else:
        for i in mediafiles:
            media_ids.append(media_upload(i))


def publish(statusfile, mediafiles):
    statuslist, end, t_num, status = load_status(statusfile)
    generate_media_ids(mediafiles)
    update_status(statuslist, media_ids, end, t_num, status)


if __name__ == '__main__':
    publish(statusfile, mediafiles)
