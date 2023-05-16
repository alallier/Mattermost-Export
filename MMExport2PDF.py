# Alexander J. Lallier
# Mattermost Export Script

import requests
import gzip
import ujson as json
import datetime
import marko
import shutil
import sys
import os
from pathlib import Path
from fpdf import FPDF, TitleStyle, HTMLMixin

# File_object = open(r"MM1.md", "w", encoding='utf8')

mattermostURL = 'https://mattermost.tld/api/v4/'
accessToken = ""
userId = ''
teamId = ''

headers = {'Authorization': 'Bearer ' + accessToken}

users = {}
channelCache = {}

baseUserPath = ''

# channelDisplayName = None
channelDisplayName = 'Table of Contents'
messageHeader = None
tableOfContents = {}

def main():
  # Get all Channels for a User
  getChannelsForAUser = f'users/{userId}/teams/{teamId}/channels?include_deleted=false&last_delete_at=0'

  allChannelsForUserResponse = requests.get(mattermostURL + getChannelsForAUser, headers=headers)

  if (allChannelsForUserResponse.status_code != 200):
    print("ERROR: Getting all channels for user")
    exit()

  allChannelsForUser = allChannelsForUserResponse.json()

  allChannelsForUser.reverse()

  hitPublicChannel = False
  hitPrivateChannel = False
  hitDMChannel = False
  hitGroupMessages = False

  # Initialize PDF
  pdf = PDF()
  pdf.add_page()
  pdf.set_auto_page_break(True, 15.0)
  # pdf.insert_toc_placeholder(render_toc, pages=1)

  publicChannels = []
  privateChannels = []
  groupChannels = []
  directMessageChannels = []

  channelGroupingsList = []

  for channel in allChannelsForUser:
    if (channel["type"] == 'O'):
      publicChannels.append(channel)

    if (channel["type"] == 'P'):
      privateChannels.append(channel)

    if (channel["type"] == 'D'):
      directMessageChannels.append(channel)

    if (channel["type"] == 'G'):
      groupChannels.append(channel)

  # Pre-process names in direct messages so we can sort by the other user's name
  for channel in directMessageChannels:
    channel['full_name'] = directMessageOtherUserName(channel)

  # Sort alphabetical
  publicChannels = sorted(publicChannels, key = lambda i: (i['name']))
  privateChannels = sorted(privateChannels, key = lambda i: (i['name']))
  groupChannels = sorted(groupChannels, key = lambda i: (i['name']))
  directMessageChannels = sorted(directMessageChannels, key = lambda i: (i['full_name']))

  channelGroupingsList = publicChannels + privateChannels + groupChannels + directMessageChannels

  for channel in channelGroupingsList:

    messagesArray = []
    pinnedMessages = []

    # Setup Channel Name and Headers for printing
    setupChannelNameAndHeader(channel)

    if (channel["type"] == 'O' and hitPublicChannel == False):
      pdf.set_fill_color(255, 165, 0)
      pdf.start_section("PUBLIC CHANNELS")
      # pdf.cell(0, 15, "PUBLIC CHANNELS", 0, 2, 'C', True)
      hitPublicChannel = True

    if (channel["type"] == 'P' and hitPrivateChannel == False):
      pdf.set_fill_color(255, 165, 0)
      pdf.start_section("PRIVATE CHANNELS")
      # pdf.cell(0, 15, "PRIVATE CHANNELS", 0, 2, 'C', True)
      hitPrivateChannel = True

    if (channel["type"] == 'D' and hitDMChannel == False):
      pdf.set_fill_color(255, 165, 0)
      pdf.start_section("DIRECT MESSAGE CHANNELS")
      # pdf.cell(0, 15, "DIRECT MESSAGE CHANNELS", 0, 2, 'C', True)
      hitDMChannel = True

    if (channel["type"] == 'G' and hitGroupMessages == False):
      pdf.set_fill_color(255, 165, 0)
      pdf.start_section("GROUP MESSAGE CHANNELS")
      # pdf.cell(0, 15, "GROUP MESSAGE CHANNELS", 0, 2, 'C', True)
      hitGroupMessages = True

    print(channelDisplayName)
    # File_object.write("## " + channelDisplayName + '\n\n')
    pdf.set_fill_color(255, 0, 0)
    pdf.start_section(channelDisplayName, level=1)
    # pdf.set_link(tableOfContents[channel["display_name"]])
    # pdf.multi_cell(0, 5, messageHeader, 0, 'L', True)
    # pdf.ln()

    channelId = channel["id"]

    morePages = True
    channelPostsCounter = 0
    allPosts = []
    allPostsFull = []
    # Get all pages and append messages to one array. We reverse this array before processing so order is from older to newest when printing
    while (morePages):
      getPostsForChannel = f'channels/{channelId}/posts?page={channelPostsCounter}'

      try:
        allPostsForChannelResponse = requests.get(mattermostURL + getPostsForChannel, headers=headers)
      except:
        print('ERROR ON GETTING PAGE')

      if (allPostsForChannelResponse.status_code != 200):
        print("ERROR: Getting all posts for channel")
        exit()

      allPostsForChannel = allPostsForChannelResponse.json()

      postFiles = []

      if not allPostsForChannel["posts"]:
        morePages = False

      channelPostsCounter += 1

      for key in allPostsForChannel["order"]:
        allPosts.append(allPostsForChannel["posts"][key])
        allPostsFull.append(allPostsForChannel)

    # CACHE CHANNEL HERE
    channelCache[channelId] = {
      "channelName": channelDisplayName,
      "posts": allPostsFull
    }

    # Reverse so it prints oldest to newest
    allPosts.reverse()

    # BEGIN POST PROCESSING
    # Loop over posts for channel
    for post in allPosts:
        pictures = []
        files = []

        message = post["message"]
        if (isinstance(message, str)):
          postUserId = post["user_id"]

          theUser = getUser(postUserId)

          # Files
          if "metadata" in post and "files" in post["metadata"]:
            postFiles = post["metadata"]["files"]

            if len(postFiles) > 0:
              for file in postFiles:
                # file["extension"] == "gif"
                if file["extension"] == "png" or file["extension"] == "jpeg" or file["extension"] == "jpg":
                  pictures.append(file)
                else:
                  files.append(file)

                # APPEND IF FOR UNIQUENESS HERE TOO
                # message = message + "\n" + "Attached file: " + file["id"] + "_" +file["name"]

          postWithUserName = {
            "name": theUser["first_name"] + " " + theUser["last_name"],
            "message": message,
            "time": str(datetime.datetime.fromtimestamp(post["create_at"] / 1000).strftime("%m/%d/%Y, %I:%M:%S %p")),
            "pictures": pictures,
            "files": files,
            "post": post
          }

          if post["is_pinned"] == True:
            pinnedMessages.append(postWithUserName)

          messagesArray.append(postWithUserName)

    print('Total Messages: ', len(messagesArray) + 1)
    print('\n')

    if len(pinnedMessages) > 0:
      # pdf.set_fill_color(255, 165, 0)
      # pdf.ln()
      pdf.start_section("Pinned Messages", level=2)
      # pdf.cell(0, 10, "Pinned Messages", 0, 1, 'C', True)
      # pdf.ln()

    # Loop through Pinned messages first, to put them all at the front
    for message in pinnedMessages:
      userName = message["name"]
      singleMessage = message["message"]
      time = message["time"]

      pdf.set_fill_color(220, 220, 220)
      pdf.set_draw_color(255, 165, 0)
      pdf.cell(0, 5, handleUnicode(userName) + " " + time + " Pinned", 0, 0, 'L', True)
      pdf.set_fill_color(255, 255, 255)
      pdf.ln()
      pdf.multi_cell(0, 5, handleUnicode(singleMessage), 1, 'L', True)
      # pdf.write_html(marko.convert(singleMessage))
      pdf.ln()

    pdf.set_draw_color(0, 0, 0)
    pdf.set_fill_color(220, 220, 220)
    pdf.start_section("Regular Messages", level=2)
    # pdf.cell(0, 10, "Regular Messages", 0, 1, 'C', True)
    pdf.set_fill_color(255, 255, 255)
    for message in messagesArray:
      userName = message["name"]
      singleMessage = message["message"]
      time = message["time"]
      post = message["post"]

      pdf.set_fill_color(220, 220, 220)
      if post["is_pinned"] == True:
        pdf.set_draw_color(255, 165, 0)
        pdf.cell(0, 5, handleUnicode(userName) + " " + time + " Pinned", 0, 0, 'L', True)

        pdf.set_fill_color(255, 255, 255)

        pdf.ln()
        pdf.multi_cell(0, 5, handleUnicode(singleMessage), 1, 'L', True)
        pdf.ln()
        pdf.set_draw_color(0, 0, 0)
      else:
        pdf.cell(0, 5, handleUnicode(userName) + " " + time, 0, 0, 'L', True)
        pdf.set_fill_color(255, 255, 255)
        pdf.ln()
        pdf.multi_cell(0, 5, handleUnicode(singleMessage), 0, 'L', True)
        pdf.ln()

      global baseUserPath
      baseUserPath = f'./users/{getUser(userId)["username"]}/'
      baseUserFilePath = baseUserPath + "files/"
      userPicturesFilePath = baseUserFilePath + "pics/"
      userAttachmentsFilePath = baseUserFilePath + "files/"

      # for picture in message["pictures"]:
      #   # APPEND FILE ID TO PATH TO MAKE UNIQUE AND CACHE THIS
      #   filePath = userPicturesFilePath + picture["id"] + "_" + picture["name"]
      #   my_file = Path(filePath)
      #   if not my_file.exists():
      #     getAFile = f'/files/{picture["id"]}'
      #     fileResponse = requests.get(mattermostURL + getAFile, headers=headers, stream=True)

      #     if (fileResponse.status_code != 200):
      #       print("ERROR: Getting a picture")
      #     else:
      #       Path(userPicturesFilePath).mkdir(parents=True, exist_ok=True)

      #       with open(filePath, 'wb') as f:
      #         fileResponse.raw.decode_content = True
      #         shutil.copyfileobj(fileResponse.raw, f)

      #         # print('Saved Picture')

      #   try:
      #     pdf.image(filePath, h = 25)
      #   except:
      #     print('Error: Couldn\'t save picture to PDF')
      #     # PRINT THE IMAGE NAME TO FILE IN CASE OF ERROR

      for aFile in message["files"]:
        filePath = userAttachmentsFilePath + aFile["id"] + '_' + aFile["name"]
        # my_file = Path(filePath)
        # if not my_file.exists():
        #   getAFile = f'/files/{aFile["id"]}'
        #   fileResponse = requests.get(mattermostURL + getAFile, headers=headers, stream=True)

        #   if (fileResponse.status_code != 200):
        #     print("ERROR: Getting a picture")
        #   else:
        #     Path(userAttachmentsFilePath).mkdir(parents=True, exist_ok=True)

        #     with open(filePath, 'wb') as f:
        #       fileResponse.raw.decode_content = True
        #       shutil.copyfileobj(fileResponse.raw, f)

        pdf.cell(30, 5, "Attached file: ", 0, 0, 'L', True)
        pdf.set_text_color(0, 0, 255)
        # pdf.cell(0, 5, file["id"] + "_" +file["name"], 0, 0, 'L', True, link="file:///./files/files/" + aFile["id"] + '_' + aFile["name"])
        pdf.cell(0, 5, file["id"] + "_" +file["name"], 0, 0, 'L', True)
        pdf.set_text_color(0, 0, 0)
        pdf.ln()
        pdf.ln()

          # print('Saved File')

  print("Writing to PDF file")
  print(baseUserPath + "messages.pdf")
  print()
  Path(baseUserPath).mkdir(parents=True, exist_ok=True)
  pdf.add_page()
  pdf.output(baseUserPath + "messages.pdf", 'F')
  makeJsonFile()

def getUser(userId):
  if userId not in users:
    getUser = f'/users/{userId}'
    getUserResponse = requests.get(mattermostURL + getUser, headers=headers)

    if (getUserResponse.status_code != 200):
      print("ERROR: Getting user")
      exit()

    user = getUserResponse.json()

    users[userId] = user

  return users[userId]

def setupChannelNameAndHeader(channel):
  global messageHeader
  global channelDisplayName

  channelDisplayName = channel["display_name"]
  # Direct messsages
  # if len(channel["display_name"]) == 0:
  if channel["type"] == 'D':
    nameSplit = channel["name"].split("__")
    firstPerson = getUser(nameSplit[0])
    firstPersonFirstName = firstPerson["first_name"]
    firstPersonLastName = firstPerson["last_name"]
    firstPersonUserId = nameSplit[0]

    secondPerson = getUser(nameSplit[1])
    secondPersonFirstName = secondPerson["first_name"]
    secondPersonLastName = secondPerson["last_name"]
    secondPersonUserId = nameSplit[1]

    if firstPersonUserId == userId:
      otherPersonFirstName = secondPersonFirstName
      otherPersonLastName = secondPersonLastName
    else:
      otherPersonFirstName = firstPersonFirstName
      otherPersonLastName = firstPersonLastName

    messageHeader = 'DM with ' + otherPersonFirstName + ' ' + otherPersonLastName
    channelDisplayName = messageHeader
  else:
    # If MM Group message
    if channel["type"] == 'G':
      # Get Channel Members:
      names = getChannelMembersFn(channel)

      messageHeader = "Group Message between: " + names
      channelDisplayName = messageHeader
    # Public/Private Channels
    else:
      messageHeader = channelDisplayName

def directMessageOtherUserName(channel):
  nameSplit = channel["name"].split("__")
  firstPerson = getUser(nameSplit[0])
  firstPersonFirstName = firstPerson["first_name"]
  firstPersonLastName = firstPerson["last_name"]
  firstPersonUserId = nameSplit[0]

  secondPerson = getUser(nameSplit[1])
  secondPersonFirstName = secondPerson["first_name"]
  secondPersonLastName = secondPerson["last_name"]
  secondPersonUserId = nameSplit[1]

  # print('First Person ID: ', firstPersonUserId)
  # print('Actual User Id:', userId)
  # print(firstPersonUserId == userId)
  if firstPersonUserId == userId:
    return secondPersonFirstName + secondPersonLastName
  else:
    return firstPersonFirstName + firstPersonLastName

def getChannelMembersFn(channel):
  channelMembersCounter = 0
  morePages = True
  names = ''
  while(morePages):
    getChannelMembers = f'/channels/{channel["id"]}/members?page={channelMembersCounter}'
    getChannelMembersResponse = requests.get(mattermostURL + getChannelMembers, headers=headers)

    if (getChannelMembersResponse.status_code != 200):
      print("ERROR: Getting all posts for channel")
      exit()

    channelMembers = getChannelMembersResponse.json()

    channelMembersCounter += 1

    channelMembersLoopCounter = 0
    for member in channelMembers:
      user = getUser(member["user_id"])

      if channelMembersLoopCounter == len(channelMembers) - 1:
        names += 'and ' + user["first_name"] + ' ' + user["last_name"]
      else :
        names += user["first_name"] + ' ' + user["last_name"] + ', '

      channelMembersLoopCounter += 1

    if len(channelMembers) == 0:
      morePages = False
      break
  return names

def handleUnicode(text):
  # newText = text.replace(u'\u2013', '-')
  newText = text.encode('latin-1', 'replace').decode('latin-1')
  return newText
  # return text

class PDF(FPDF, HTMLMixin):
  def __init__(self):
    super().__init__()

    SYSTEM_TTFONTS = './fonts'

    self.add_font("NotoSans", style="", fname=SYSTEM_TTFONTS + "/NotoSans-Regular.ttf", uni=True)
    self.add_font("NotoSans", style="B", fname=SYSTEM_TTFONTS + "/NotoSans-Bold.ttf", uni=True)
    self.add_font("NotoSans", style="I", fname=SYSTEM_TTFONTS + "/NotoSans-Italic.ttf", uni=True)
    self.add_font("NotoSans", style="BI", fname=SYSTEM_TTFONTS + "/NotoSans-BoldItalic.ttf", uni=True)
    # self.add_font('DejaVu', '', 'DejaVuSansCondensed.ttf', uni=True)
    self.set_font('NotoSans', '', 10)

    self.set_section_title_styles(
      # Level 0 titles:
      TitleStyle(
          font_family="Times",
          font_style="B",
          font_size_pt=24,
          color=(0,0,0),
          underline=True,
          t_margin=5,
          l_margin=0,
          b_margin=5,
      ),
      # Level 1 subtitles:
      TitleStyle(
          font_family="Times",
          font_style="B",
          font_size_pt=20,
          color=(0,0,0),
          underline=True,
          t_margin=5,
          l_margin=0,
          b_margin=5,
      ),
      # Level 2 subtitles:
      TitleStyle(
          font_family="Times",
          font_style="B",
          font_size_pt=15,
          color=(255, 165, 0),
          underline=True,
          t_margin=5,
          l_margin=0,
          b_margin=5,
      )
    )

  def header(self):
      # Select Arial bold 15
      self.set_font("NotoSans", style='B', size=12)
      # Move to the right
      self.cell(80)
      # Framed title
      self.cell(30, 10, channelDisplayName, 0, 2, 'C')
      # Line break
      self.ln(15)
  def footer(self):
      # Go to 1.5 cm from bottom
      self.set_y(-15)
      # Select Arial italic 8
      self.set_font("NotoSans", style='I', size=8)
      # Print centered page number
      self.cell(0, 10, 'Page %s' % self.page_no(), 0, 2, 'C')

def render_toc(pdf, outline):
  pdf.y += 10
  # pdf.set_font("Helvetica", size=16)
  pdf.underline = True
  p(pdf, "Table of contents:")
  pdf.underline = False
  pdf.y += 5
  # pdf.set_font("Courier", size=12)
  for section in outline:
      print(section)
      link = pdf.add_link()
      pdf.set_link(link, page=section.page_number)
      text = f'{" " * section.level * 4} {section.name}'
      text += (
          f' {"." * (60 - section.level*2 - len(section.name))} {section.page_number}'
      )
      pdf.multi_cell(w=pdf.epw, h=pdf.font_size, txt=text, ln=1, align="L", link=link)

def p(pdf, text, **kwargs):
    pdf.multi_cell(w=pdf.epw, h=pdf.font_size, txt=text, ln=1, **kwargs)

def makeJsonFile():
  ## PRINT STATEMENT FOR JSON FILE NEEDED
  filePath = baseUserPath + "messages.gz"
  print("Writing JSON to file")
  print(filePath)
  with gzip.open(filePath, 'wt', encoding="ascii") as zipfile:
    json.dump(channelCache, zipfile)
  # with open(filePath, "w+") as outfile:
  #   json.dump(channelCache, outfile)

main()
