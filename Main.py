import xml
import xml.dom.minidom
import urllib.request
import os
import eyed3.id3
from mutagen.easyid3 import EasyID3
from mutagen.id3 import ID3, TPE1, ID3NoHeaderError
import time
from optparse import OptionParser
import gspread
from oauth2client.service_account import ServiceAccountCredentials


def optionParser():
    parser = OptionParser()
    parser.add_option("-f", "--file", dest="path",
                      help="Specify Podcast Root Directory", metavar="FILE")
    parser.add_option("-c", "--credentials", dest="cred",
                      help="Specify Sheets Credentials", metavar="CRED")
    (options, args) = parser.parse_args()
    return options, args


def sheetsAuth():
    scope = ['https://spreadsheets.google.com/feeds']
    auth = ServiceAccountCredentials.from_json_keyfile_name(sheets_Cred, scope)
    client = gspread.authorize(auth)
    sheet = client.open("Podcast List").sheet1
    podcast_list = sheet.col_values(4)
    return (podcast_list, sheet)


def parseXML(xml_site):
    xml_string = urllib.request.urlopen(xml_site).read()
    xml_parsed = xml.dom.minidom.parseString(xml_string)
    return xml_parsed


def getFirstChannel(xml_parsed):
    Channel = xml_parsed.getElementsByTagName("channel")[0]
    return Channel


def getFirstItem(xml_parsed):
    Item = xml_parsed.getElementsByTagName("item")[0]
    return Item


def returnItemElement(ElementString, Item):
    nodeElement = Item.getElementsByTagName(ElementString)
    elementText = nodeElement.item(0)
    elementFirstChild = elementText.firstChild
    Element = elementFirstChild.data
    return Element


def returnChannelElement(ElementString, Channel):
    elementText = Channel.getElementsByTagName(ElementString)[0]
    elementFirstChild = elementText.firstChild
    Element = elementFirstChild.data
    return Element


def returnImageElement(ElementString):
    if Channel.getElementsByTagName("image").length == 0:
        Element = getURL(Channel, "itunes:image", 1)
    else:
        Image = Channel.getElementsByTagName("image")[0]
        elementText = Image.getElementsByTagName(ElementString)[0]
        elementFirstChild = elementText.firstChild
        Element = elementFirstChild.data
    return Element


def getURL(item, urlstring, parse):
    URL1 = item.getElementsByTagName(urlstring)[0]
    URL2 = URL1.toxml()
    URL = URL2.split("\"")[parse]
    # URL=returnElement("enclosure",item)
    return URL


def id3Tagging():
    podcast_file = eyed3.load(directory)
    podcast_file.tag.artist = pod_author
    podcast_file.tag.album = pod_title
    podcast_file.tag.title = episode_title
    podcast_file.tag.date = pubDate
    podcast_file.tag.comment = episode_comment
    podcast_file.tag.genre = 'Podcast'
    podcast_file.tag.save()


def mutagenid3Tagging():
    try:
        podcast_file = EasyID3(directory)
    except ID3NoHeaderError:
        podcast_file = ID3()
        podcast_file.add(TPE1(encoding=3, text=u'Artist'))
        podcast_file.add(TPE1(encoding=3, text=u'Album'))
        podcast_file.add(TPE1(encoding=3, text=u'Title'))
        podcast_file.add(TPE1(encoding=3, text=u'Date'))
        podcast_file.add(TPE1(encoding=3, text=u'Genre'))
        podcast_file.save(directory)
    podcast_file = EasyID3(directory)
    podcast_file['Artist'] = pod_author
    podcast_file['Album'] = pod_title
    podcast_file['Title'] = episode_title
    podcast_file['Date'] = pubDate
    podcast_file['Genre'] = 'Podcast'
    podcast_file.save()


def remove(path):
    """
    Remove the file or directory
    """
    if os.path.isdir(path):
        try:
            os.rmdir(path)
        except OSError:
            print("Unable to remove folder: %s" % path)
    else:
        try:
            if os.path.exists(path):
                os.remove(path)
        except OSError:
            print("Unable to remove file: %s" % path)


def cleanup(number_of_days, path):
    """
    Removes files from the passed in path that are older than or equal
    to the number_of_days
    """
    time_in_secs = time.time() - (number_of_days * 24 * 60 * 60)
    for root, dirs, files in os.walk(path, topdown=False):
        for file_ in files:
            full_path = os.path.join(root, file_)
            stat = os.stat(full_path)
            if (stat.st_mtime <= time_in_secs) & (".mp3" in full_path):
                remove(full_path)

        if not os.listdir(root):
            remove(root)


opener = urllib.request.build_opener()
opener.addheaders = [('User-Agent',
                      'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/36.0.1941.0 Safari/537.36')]
urllib.request.install_opener(opener)
(option, args) = optionParser()
download_Directory = option.path
sheets_Cred = option.cred
(pod_List, sheet) = sheetsAuth()
i = 1

for xml_site in pod_List:
    if (str(xml_site) == 'Podcast URL') | (str(xml_site) == ''):
        i = i + 1
        continue
    xml_parsed = parseXML(xml_site)
    Item = getFirstItem(xml_parsed)
    Channel = getFirstChannel(xml_parsed)
    episode_comment = returnItemElement("description", Item)
    pod_title = returnChannelElement("title", Channel)
    episode_title = returnItemElement("title", Item)
    episode_title = episode_title.replace("/", "")
    pod_author = returnChannelElement("itunes:author", Channel)
    pod_author = pod_author.split("/")[0]
    download_path = str(download_Directory + pod_author + "/" + pod_title + "/")
    directory = os.path.dirname(download_path)
    if not os.path.exists(directory):
        print(time.strftime('%Y-%m-%d %H:%M:%S') + " Adding New Podcast: " + pod_title)
        sheet.update_cell(i, 1, pod_title)
        os.makedirs(directory)
        Image_URL = returnImageElement('url')
        urllib.request.urlretrieve(Image_URL, directory + "/cover.jpg")
    directory = os.path.dirname(download_path) + "/" + episode_title + ".mp3"
    if os.path.exists(directory):
        i=i+1
        continue
    pubDate = returnItemElement("pubDate", Item).split("+")[0]
    url = getURL(Item, "enclosure", 5)
    print(time.strftime('%Y-%m-%d %H:%M:%S') + " Downloading new " + pod_title + " episode " + episode_title)
    print("")
    try:
        urllib.request.urlretrieve(url, directory)
        sheet.update_cell(i, 3, time.strftime('%Y-%m-%d %H:%M:%S'))
    except "HTTP Error 403":
        sheet.update_cell(i, 3, "Error")
        continue

    sheet.update_cell(i, 2, episode_title)
    urllib.request.urlcleanup()
    i=i+1
    cleanup(7, os.path.dirname(download_path))
    try:
        mutagenid3Tagging()
    except:
        try:
            id3Tagging()
        except:
            print('Error-ID3 Tag Failed to Save for:')
            print(pod_title)
            print(episode_title)
            print(pod_author)
            print(pubDate)
            print(episode_comment)
            continue
