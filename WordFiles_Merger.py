# -*- coding: UTF-8 -*-
#!/usr/bin/env python
#-------------------------------------------------------------------------------
# Name :       Word Joiner
# Purpose :    Generate a Word file compiling all word files present in folder
#				and subfolders according to a prefix filter.
# Authors :    Julien M.
# Python :     2.7.x
# Encoding:    utf-8
# Created :    14/03/2013
# Updated :    01/04/2013
# Version :    0.1
#-------------------------------------------------------------------------------

###################################
####### Modules importation #######
###################################

# Standard library
from os import walk, path
from sys import platform
from sys import exit
from Tkinter import Tk
from tkFileDialog import askdirectory

# External library
from win32com.client import *

###################################
########### Functions #############
###################################

def listword(foldpath, prefix = '*'):
    u""" List Word files included in the folder and its subfolders """
    global wordfiles
    extensions = ['.doc', '.docx']  # Word files extensions
    # Looping the folders structure
    for root, dirs, files in walk(target):
        for f in files:
            if path.splitext(f)[1] in extensions and prefix in f:
                wordfiles.append(path.normpath(path.join(root, f)))
    # Sorting and tupling
    wordfiles.sort()
    wordfiles = tuple(wordfiles)
    # End of function
    return wordfiles

def mergeword(iterword, dest, output):
    u""" create a new Word file (.doc/.docx) merging all others Word files
    contained into the iterable parameter (list or tuple) """
    # Initializing Word application
    word = gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    # Create the final document
    finaldoc = word.Documents.Add()
    # Looping and merging
    for f in iterword:
        rng = finaldoc.Range()
        rng.Collapse(0)
        rng.InsertFile(f)
        rng.Paragraphs.Add()
        rng.Collapse(0)
        rng.InsertBreak()
        del rng
    # Page numbers
    sec = finaldoc.Sections.Item(1)
    bdp = sec.Footers.Item(1)
    bdp.PageNumbers.Add()
    # saving
    finaldoc.SaveAs(path.join(dest, output + '.doc'), FileFormat=0)
    # Trying to convert into newer version of Office
    try:
        """ Office newer than 2003 is installed """
        finaldoc.Convert()
    except:
        """ Just Office 2003 """
        None
    # clean up
    finaldoc.Close()
    word.Quit()
    # end of function
    return finaldoc

###################################
######## Global variables #########
###################################

wordfiles = []  # list of Word files

###################################
########## Main program ###########
###################################

# Check if it's running on windows system
if platform != 'win32':
	print u"Sorry, it's only working for Windows operating system!"
	exit()

# Ask for the "folder-target"
root = Tk()
root.withdraw()
target = askdirectory(mustexist = True)         # GUI for choose folder

if target == "":          # if operation cancelled, stop the machine
    root.destroy()
    exit()

# Ask for prefix filter and name of output file
prefix = raw_input("Prefix for filter Word files: ")
outname = raw_input("Name of output file: ")

# List Word files contained
listword(target, prefix)

# Merge all files
mergeword(wordfiles, target, outname)

# Closing
exit()