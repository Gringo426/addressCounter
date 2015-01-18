'''
File: addressCounter.py
Author: Ryan Grings
Version: 1.0
Description: Determines amount of emails from each unique address in the Windows Live Mail Inbox folder.
'''

import glob
import os
import re

def emailCount():
    userName = os.getenv("USERNAME")
    startDir = "C:\\Users\\"+userName+"\\AppData\\Local\\Microsoft\\Windows Live Mail\\"
    try:
        os.chdir(startDir)
    except NotADirectoryError or FileNotFoundError:
        quit(0)
    for root, dirs, files in os.walk(os.getcwd()):
        if "Inbox" in dirs:
            os.chdir(root+"\\Inbox")
            listEmails()
        else:
            continue
        
def listEmails():
    listEmail = glob.glob("*.eml")
    patternOne = re.compile("From: .*<.+>")
    patternTwo = re.compile("From: .+@.+\S{3}")
    patternThree = re.compile("From: .*<.+>", re.DOTALL)
    addressList = []
    decodeError = 0
    noSender = 0
    for fileName in listEmail:
        inputFile = open(fileName, "r")
        try:
            tempInput = inputFile.read()
            matchOne = patternOne.search(tempInput)
            if matchOne:
                addressList.append(matchOne.group().split("<")[1].split(">")[0])
            else:
                matchTwo = patternTwo.search(tempInput)
                if matchTwo:
                    addressList.append(matchTwo.group().split(" ")[-1])
                else:
                	matchThree = patternThree.search(tempInput)
                	if matchThree:
                		addressList.append(matchThree.group().split("<")[1].split(">")[0])
                	else:
                		print(fileName)
                		noSender+=1
        except UnicodeDecodeError:
            decodeError+=1
            continue
    for element in set(addressList):
        print(element, addressList.count(element), sep=': ')
    if decodeError > 0:
    	print("Unable to decode", decodeError, "emails.")
    if noSender > 0:
    	print("Unable to find sender for", noSender, "emails.")
    print("Total emails: " + str(len(addressList)+noSender+decodeError))
    print("-------------------------------------------------------------")

    return

if __name__ == "__main__":
    emailCount()
