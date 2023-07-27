#!/usr/bin/env python3
# -*- coding: utf-8 -*-

''' MMExport2PDF

Using the Mattermost API, connects to an instance and exports
all channel for a user on a team.

Images and Files are downloaded as well.

'''

#########################
## Python Imports
##

import argparse
import requests
import gzip
import simplejson as json
#import ujson as json
import datetime
import shutil
import sys
import os
from pathlib import Path

#import traceback

#########################
## Thirdparty Imports
##

# fpdf is acutally PyFPDF2
from fpdf import FPDF, TitleStyle, Align
from fpdf.enums import FileAttachmentAnnotationName 

__author__ = 'Alexander J. Lallier'
__version__ = '1.0'
__contact__ = ''



#########################
## Globals Variables
##

imageExtenstions = [ 'gif', 'png', 'jpeg', 'jpg' ]

mattermostURL = ''
headers = {}
baseUserPath = ''

users = {}
channelCache = {}


channelDisplayName = ''
messageHeader = None
tableOfContents = {}



#########################
## Exception Definitions
##

class OptionsException( Exception ):
    def __init__(self, message = None ):
        super(OptionsException,self).__init__(message)

class UserInfoException( Exception ):
    def __init__(self, message = None ):
        super(UserInfoException,self).__init__(message)

class UserIDException( Exception ):
    def __init__(self, message = None ):
        super(UserIDException,self).__init__(message)

class TeamIDException( Exception ):
    def __init__(self, message = None ):
        super(TeamIDException,self).__init__(message)

class UserChannelsException( Exception ):
    def __init__(self, message = None ):
        super(UserChannelsException,self).__init__(message)

class ImageException( Exception ):
    def __init__(self, message = None ):
        super(ImageException,self).__init__(message)

class FileException( Exception ):
    def __init__(self, message = None ):
        super(FileException,self).__init__(message)

class ChannelPostsException( Exception ):
    def __init__(self, message = None ):
        super(ChannelPostsException,self).__init__(message)

class ChannelMembersException( Exception ):
    def __init__(self, message = None ):
        super(ChannelMembersException,self).__init__(message)


#########################
## MMExport2PDF Options
##

def processOptions():
    '''
    Process command line arguments and set the internal options appropriately.

            @param argv List of command line arguments.
            @return The object containing the processed options.
    '''
    # process options

    options = None

    try:
        usage = f'%(prog)s [options]'
        description = '%(prog)s is used to export all a users channels and DMs from a team.'
        epilog = 'This can take a long time to run.'

        parser = argparse.ArgumentParser(usage=usage,
                                         description=description,
                                         epilog=epilog,
                                         formatter_class=argparse.ArgumentDefaultsHelpFormatter)

        usergroup = parser.add_argument_group(title='User Info')
        usergroup.add_argument("-a", "--auth", help="Auth Token", action="store", dest="auth", required=True)
        usergroup.add_argument("-u", "--user", help="Username of user to be exported", action="store", dest="user", required=True)
        usergroup.add_argument("-t", "--team", help="Team to export from", action="store", dest="team", required=True)

        servergroup = parser.add_argument_group(title='Server Info')
        servergroup.add_argument("-s", "--server", help="Hostname or IP of the server", action="store", dest="server", default="mattermost.com")

        categorygroup = parser.add_argument_group(title='Channel Categories')
        categorygroup.add_argument("-p", "--public", help="Exclude public channels", action="store_true", dest="public")
        categorygroup.add_argument("-P", "--private", help="Exclude private channels", action="store_true", dest="private")
        categorygroup.add_argument("-g", "--groups", help="Exclude group messages", action="store_true", dest="group")
        categorygroup.add_argument("-d", "--DMs", help="Exclude direct messages", action="store_true", dest="dms")

        filtergroup = parser.add_argument_group(title='Message Filters')
        filtergroup.add_argument("-I", "--include", help="Only inlcude these channels in the export.", nargs='*', dest="include", default=[])
        filtergroup.add_argument("-E", "--exclude", help="Exclude these channels from the export", nargs='*', dest="exclude", default=[])

        exportgroup = parser.add_argument_group(title='Export Options')
        exportgroup.add_argument("-i", "--images", help="Embed images in PDF", action="store_true", dest="images")
        exportgroup.add_argument("-f", "--files", help="Embed files in PDF", action="store_true", dest="files")
        exportgroup.add_argument("-j", "--json", help="Export JSON", action="store_true", dest="json")
        exportgroup.add_argument("-o", "--output", help="Base output directory", action="store", dest="output", default='./users')

        options = parser.parse_args() # uses sys.argv[1:] by default


    except Exception as e: #pylint: disable=broad-except
        raise OptionsException( e )

    return options


#########################
## Main
##

def main():

    try:
        global baseUserPath
        global mattermostURL
        global headers

        options = processOptions()

        if (options.public and options.private and options.group and options.dms):
            raise OptionsException( 'At least one channel category must be exported' )

        # Setup
        mattermostURL = f'https://{options.server}/api/v4/'
        headers['Authorization'] = f'Bearer {options.auth}'

        userInfo = getUserFromName(options.user)
        teamInfo = getTeam(options.team)

        baseUserPath = os.path.join( options.output, options.user )
        baseUserFilePath = os.path.join( baseUserPath, 'files/' )

        os.makedirs( baseUserPath, 0o755, True)

        # Start Working
        allChannelsForUser = getChannelsForAUser(userInfo['id'], teamInfo['id'])
        allChannelsForUser.reverse()


        hitPublicChannel = False
        hitPrivateChannel = False
        hitDMChannel = False
        hitGroupMessages = False

        # Initialize PDF
        pdf = PDF()
        pdf.add_page()
        pdf.set_auto_page_break(True, 15.0)

        publicChannels = []
        privateChannels = []
        groupChannels = []
        directMessageChannels = []

        channelGroupingsList = []
        
        for channel in allChannelsForUser:            
            if ( channel["display_name"] not in options.exclude):
                if ( (not options.include) or (channel["display_name"] in options.include) ):
                    if ((not options.public) and channel["type"] == 'O'):
                        publicChannels.append(channel)

                    if ((not options.private) and channel["type"] == 'P'):
                        privateChannels.append(channel)

                    if ((not options.dms) and channel["type"] == 'D'):
                        directMessageChannels.append(channel)

                    if ((not options.group) and channel["type"] == 'G'):
                        groupChannels.append(channel)

        # Pre-process names in direct messages so we can sort by the other user's name
        for channel in directMessageChannels:
            channel['full_name'] = directMessageOtherUserName(channel, userInfo['id'])

        # Sort alphabetical

        publicChannels = sorted(publicChannels, key = lambda i: (i['name']))
        privateChannels = sorted(privateChannels, key = lambda i: (i['name']))
        groupChannels = sorted(groupChannels, key = lambda i: (i['name']))
        directMessageChannels = sorted(directMessageChannels, key = lambda i: (i['full_name']))

        channelGroupingsList = publicChannels + privateChannels + groupChannels + directMessageChannels
        
        if not channelGroupingsList:
            raise ChannelPostsException( "No posts matched the export criteria" )
        
        for channel in channelGroupingsList:

            messagesArray = []
            pinnedMessages = []

            # Setup Channel Name and Headers for printing
            setupChannelNameAndHeader(channel, userInfo['id'])

            if (channel["type"] == 'O' and hitPublicChannel == False):
                pdf.set_fill_color(255, 165, 0)
                pdf.start_section("PUBLIC CHANNELS")
                hitPublicChannel = True

            if (channel["type"] == 'P' and hitPrivateChannel == False):
                pdf.set_fill_color(255, 165, 0)
                pdf.start_section("PRIVATE CHANNELS")
                hitPrivateChannel = True

            if (channel["type"] == 'D' and hitDMChannel == False):
                pdf.set_fill_color(255, 165, 0)
                pdf.start_section("DIRECT MESSAGE CHANNELS")
                hitDMChannel = True

            if (channel["type"] == 'G' and hitGroupMessages == False):
                pdf.set_fill_color(255, 165, 0)
                pdf.start_section("GROUP MESSAGE CHANNELS")
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
            # Get all pages and append messages to one array.
            # We reverse this array before processing so order is from older to newest when printing

            while (morePages):

                allPostsForChannel = getPostsForChannel(channelId, channelPostsCounter)

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
                                if file["extension"].lower() in imageExtenstions:
                                    pictures.append(file)
                                else:
                                    files.append(file)

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
                pdf.start_section("Pinned Messages", level=2)

            # Loop through Pinned messages first, to put them all at the front
            for message in pinnedMessages:
                userName = message["name"]
                singleMessage = message["message"]
                time = message["time"]

                #pdf.set_fill_color(220, 220, 220)
                pdf.set_fill_color(255, 165, 0)
                pdf.set_draw_color(255, 165, 0)
                pdf.cell(0, 5, f'{handleUnicode(userName)} {time} Pinned', 0, align='L', fill=True)
                pdf.set_fill_color(255, 255, 255)
                pdf.ln()
                pdf.multi_cell(0, 5, handleUnicode(singleMessage), 1, align='L', fill=True, markdown=True)
                # pdf.write_html(marko.convert(singleMessage))
                pdf.ln()

            pdf.set_draw_color(0, 0, 0)
            pdf.set_fill_color(220, 220, 220)
            pdf.start_section("Regular Messages", level=2)

            pdf.set_fill_color(255, 255, 255)
            for message in messagesArray:
                userName = message["name"]
                singleMessage = message["message"]
                time = message["time"]
                post = message["post"]

                if post["is_pinned"] == True:
                    pdf.set_fill_color(255, 165, 0)
                    pdf.set_draw_color(255, 165, 0)
                    pdf.cell(0, 5, f'{handleUnicode(userName)} {time} Pinned', 0, align='L', fill=True)
                    pdf.set_fill_color(255, 255, 255)

                    pdf.ln()
                    pdf.multi_cell(0, 5, handleUnicode(singleMessage), 1, align='L', fill=True, markdown=True)
                    pdf.ln()
                    pdf.set_draw_color(0, 0, 0)
                else:
                    pdf.set_fill_color(220, 220, 220)
                    pdf.cell(0, 5, f'{handleUnicode(userName)} {time}', 0, align='L', fill=True)
                    pdf.set_fill_color(255, 255, 255)
                    pdf.ln()
                    pdf.multi_cell(0, 5, handleUnicode(singleMessage), 0, align='L', fill=True, markdown=True)
                    pdf.ln()


                if( options.images ):
                    try:
                        userPicturesFilePath = os.path.join( baseUserFilePath, "pics/" )
                        os.makedirs( userPicturesFilePath, 0o755, True)

                        for picture in message["pictures"]:
                            try:
                                # APPEND FILE ID TO PATH TO MAKE UNIQUE AND CACHE THIS
                                imagePath = os.path.join( userPicturesFilePath,  f'{picture["id"]}_{picture["name"]}' )
                                myImage = Path(imagePath)

                                if not myImage.exists():
                                    imageObj = getFile( picture["id"] )

                                    with open(imagePath, 'wb') as f:
                                        imageObj.raw.decode_content = True
                                        shutil.copyfileobj(imageObj.raw, f)

                                pdf.image(imagePath, w=(pdf.epw * .75), x=Align.C)

                            except ImageException as ie:
                                print( f'Embed Image error: {ie}' )
                                #traceback.print_exc()
                            except Exception as e:
                                print('Embed Image error: Couldn\'t add picture to PDF')
                                print( e )
                                #traceback.print_exc()

                    except ImageException as ie:
                        print( ie )

                if( options.files ):
                    try:
                        userAttachmentsFilePath = os.path.join( baseUserFilePath, "files/" )
                        os.makedirs( userAttachmentsFilePath, 0o755, True)

                        for aFile in message["files"]:
                            try:
                                filePath = os.path.join( userAttachmentsFilePath, f'{aFile["id"]}_{aFile["name"]}' )
                                myFile = Path(filePath)

                                if not myFile.exists():
                                    fileObj = getFile( aFile["id"] )

                                    with open(filePath, 'wb') as f:
                                        fileObj.raw.decode_content = True
                                        shutil.copyfileobj(fileObj.raw, f)
                                                                
                                if myFile.is_file():                                    
                                    pdf.embed_file( myFile, desc=aFile["name"], compress=True)
                                    pdf.cell(30, 5, 'Attached file: ', 0, align='L', fill=True)
                                    pdf.set_text_color(0, 0, 255)
                                    pdf.cell(0, 5, f'{aFile["id"]}_{aFile["name"]}', 0, align='L', fill=True)
                                    
                            except FileException as fe:
                                print( f'Embed File error: {fe}' )
                                #traceback.print_exc()
                            except Exception as e:
                                print('Embed File error: Couldn\'t add file to PDF')
                                print( e )
                                #traceback.print_exc()
                            finally:
                                pdf.set_text_color(0, 0, 0)
                                pdf.ln()
                                
                    except ImageException as ie:
                        print( ie )

        pdfOutput = os.path.join(baseUserPath, f'{options.user}.pdf' )

        print( pdfOutput )
        print()
        pdf.add_page()
        pdf.output( pdfOutput )

        if( options.json ):
            makeJsonFile(options.user)

    except Exception as e:
        print( e )
        #traceback.print_exc()



#########################
## Helper Functions
##

def getUser(userID):
    '''
    getUser

    Returns the user info for the given ID.

        @param userID The user ID to look up

    :raises:
        UserInfoException
    '''
    if userID not in users:
        getUserResponse = requests.get(f'{mattermostURL}/users/{userID}',
                                       headers=headers)

        if (getUserResponse.status_code != 200):
            raise UserInfoException(f'Failed to get user info for: {userID}')

        users[userID] = getUserResponse.json()

    return users[userID]


def getUserFromName(username):
    '''
    getUserFromName

    Retrieves the user info for the username.

        @param username the username to look up.

    :raises:
        UserIDException
    '''
    getUserIDResponse = requests.get(f'{mattermostURL}/users/username/{username}',
                                     headers=headers)

    if (getUserIDResponse.status_code != 200):
      raise UserIDException(f'Failed to get user ID for: {username}')

    return getUserIDResponse.json()


def getTeam(team):
    '''
    getTeadID

    Returns the ID for the team.

        @param team the team name to look up.

    :raises:
        TeamIDException
    '''

    getTeamIDResponse = requests.get(f'{mattermostURL}/teams/name/{team}',
                                     headers=headers)

    if (getTeamIDResponse.status_code != 200):
      raise TeamIDException(f'Failed to get user ID for: {username}')

    return getTeamIDResponse.json()


def getFile(fileID):
    '''
    getFile

    Retrieves an attachement file from the server.

        @param fileID the attachment ID/

    :raises:
        FileException
    '''

    getFileResponse = requests.get(f'{mattermostURL}/files/{fileID}',
                                   headers=headers,
                                   stream=True)

    if (getFileResponse.status_code != 200):
      raise FileException(f'Failed to get file[{fileID}], status code: {getFileResponse.status_code}')

    return getFileResponse


def getChannelsForAUser(userID, teamID):
    '''
    getChannelsForAUser

    Get all Channels for a User

        @param userID
        @param teamID

    :raises:
        UserChannelsException
    '''
    allChannelsForUserResponse = requests.get(f'{mattermostURL}/users/{userID}/teams/{teamID}/channels?include_deleted=false&last_delete_at=0',
                                              headers=headers)

    if (allChannelsForUserResponse.status_code != 200):
        raise UserChannelsException('Failed to get channels for user')

    return allChannelsForUserResponse.json()


def getPostsForChannel(channelID, channelPostsCounter):
    '''
    getPostsForChannel

    Get all Posts for a Channels

        @param channelID
        @param channelPostsCounter

    :raises:
        ChannelPostsException
    '''
    getPostsForChannelResponse = requests.get(f'{mattermostURL}channels/{channelID}/posts?page={channelPostsCounter}', headers=headers)

    if (getPostsForChannelResponse.status_code != 200):
        raise ChannelPostsException('Failed to get posts for channels')

    return getPostsForChannelResponse.json()


def setupChannelNameAndHeader(channel, userID):
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

        if firstPersonUserId == userID:
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


def directMessageOtherUserName(channel, userID):
    nameSplit = channel["name"].split("__")
    firstPerson = getUser(nameSplit[0])
    firstPersonFirstName = firstPerson["first_name"]
    firstPersonLastName = firstPerson["last_name"]
    firstPersonUserId = nameSplit[0]

    secondPerson = getUser(nameSplit[1])
    secondPersonFirstName = secondPerson["first_name"]
    secondPersonLastName = secondPerson["last_name"]
    secondPersonUserId = nameSplit[1]

    if firstPersonUserId == userID:
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
            raise ChannelPostsException("ERROR: Getting all posts for channel")

        channelMembers = getChannelMembersResponse.json()

        channelMembersCounter += 1

        channelMembersLoopCounter = 0
        for member in channelMembers:
            user = getUser(member["user_id"])

            if channelMembersLoopCounter == len(channelMembers) - 1:
                names += 'and ' + user["first_name"] + ' ' + user["last_name"]
            else:
                names += user["first_name"] + ' ' + user["last_name"] + ', '

            channelMembersLoopCounter += 1

        if len(channelMembers) == 0:
            morePages = False
            break
    return names



def handleUnicode(text):
    newText = text.encode('latin-1', 'replace').decode('latin-1')
    return newText



class PDF(FPDF):
    def __init__(self):
        super().__init__()

        SYSTEM_TTFONTS = '/usr/share/fonts/truetype'

        self.add_font("NotoSans", style="", fname=os.path.join(SYSTEM_TTFONTS, "noto/NotoSans-Regular.ttf"))
        self.add_font("NotoSans", style="B", fname=os.path.join(SYSTEM_TTFONTS, "noto/NotoSans-Bold.ttf"))
        self.add_font("NotoSans", style="I", fname=os.path.join(SYSTEM_TTFONTS, "noto/NotoSans-Italic.ttf"))
        self.add_font("NotoSans", style="BI", fname=os.path.join(SYSTEM_TTFONTS, "noto/NotoSans-BoldItalic.ttf"))
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

        if( channelDisplayName ):
            self.multi_cell(w=0, txt=channelDisplayName, align='C')

        # Line break
        self.ln(15)


    def footer(self):
        # Go to 1.5 cm from bottom
        self.set_y(-15)
        # Select Arial italic 8
        self.set_font("NotoSans", style='I', size=8)
        # Print centered85 page number
        self.cell(0, 10, f'Page {self.page_no()}', 0, align='C')


def makeJsonFile(username):
    '''
    makeJsonFile

    Export the messages as JSON

        @param username

    '''
    ## PRINT STATEMENT FOR JSON FILE NEEDED
    jsonPath = os.path.join( baseUserPath, f'{username}.gz' )
    print("Writing JSON to file")
    print(jsonPath)
    with gzip.open(jsonPath, 'wt', encoding="ascii") as zipfile:
        json.dump(channelCache, zipfile)


if __name__ == '__main__':
  main()
