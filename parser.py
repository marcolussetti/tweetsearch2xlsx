#! python3

# Take a saved twitter search (HTML file) and return the data as a MS Excel file (xlsx file)
# Copyright (c) 2016 Marco Lussetti.

import sys
from datetime import datetime

import lxml.html
import xlsxwriter

# If no arguments provided, ask user for input

if len(sys.argv) == 3:
    input_path = sys.argv[1]
    output_path = sys.argv[2]
else:
    input_path = input("Enter input file with path: ")
    output_path = input("Enter output file with path: ")

# Open file
myfile = open(input_path, encoding="Latin-1")
myfileAsAString = myfile.read()
tree = lxml.html.fromstring(myfileAsAString)

# Extract the list of dates
extracted_dates = tree.xpath('//*[@class[starts-with(., "tweet-timestamp js-permalink js-nav js-tooltip")]]')
dates = []
dateFormat = "%I:%M %p - %d %b %Y"
# Extract dates to an array of dateformat objects
for date in extracted_dates:
    if 'title' in date.attrib:
        if "-" in date.attrib['title']:
            dates.append(datetime.strptime(date.attrib['title'], dateFormat))

# Extract list of usernames
extracted_usernames = tree.xpath('//span[@class="username js-action-profile-name"]/b')
usernames = []
for username in extracted_usernames:
    usernames.append(username.text)

# Extract list of retweets
extracted_retweets = tree.xpath(
    '//*[@class="ProfileTweet-action--retweet u-hiddenVisually"]/span[@class="ProfileTweet-actionCount"]/span[@class="ProfileTweet-actionCountForAria"]')
retweets = []
for retweet in extracted_retweets:
    if "retweets" in retweet.text:
        temp = int(retweet.text[:-8])
        retweets.append(temp)
    elif " retweet" in retweet.text:
        temp = int(retweet.text[:-7])
        retweets.append(temp)

# Extract list of likes
extracted_likes = tree.xpath(
    '//*[@class="ProfileTweet-action--favorite u-hiddenVisually"]/span[@class="ProfileTweet-actionCount"]/span[@class="ProfileTweet-actionCountForAria"]')
likes = []
for like in extracted_likes:
    if " likes" in like.text:
        temp = int(like.text[:-6])
        likes.append(temp)
    elif " like" in like.text:
        temp = int(like.text[:-5])
        likes.append(temp)

# Extract messages
extracted_messages = tree.xpath('//p[@class="TweetTextSize  js-tweet-text tweet-text"]')
messages = []
for message in extracted_messages:
    messages.append(message.text_content())

# Create xlsx file
workbook = xlsxwriter.Workbook(output_path)
worksheet = workbook.add_worksheet()

# Dateformat and headers
date_format = workbook.add_format({'num_format': 'd mmmm yyyy'})
bold = workbook.add_format({'bold': True})
worksheet.write_string(0, 0, "Date", bold)
worksheet.write_string(0, 1, "Username", bold)
worksheet.write_string(0, 2, "Message", bold)
worksheet.write_string(0, 3, "Retweets", bold)
worksheet.write_string(0, 4, "Likes", bold)

# Write lines
row = 1
for date in dates:
    worksheet.write_datetime(row, 0, date, date_format)
    row += 1

row = 1
for username in usernames:
    worksheet.write_string(row, 1, username)
    row += 1

row = 1
for message in messages:
    worksheet.write_string(row, 2, message)
    row += 1

row = 1
for retweet in retweets:
    worksheet.write_number(row, 3, retweet)
    row += 1

row = 1
for like in likes:
    worksheet.write_number(row, 4, like)
    row += 1

workbook.close()
