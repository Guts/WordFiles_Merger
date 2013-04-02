import win32com.client

def addText(text, bookmark=None):
    """ http://mail.python.org/pipermail/tutor/2002-January/011392.html """
    global worddoc, wordapp
    if bookmark:
        worddoc.Bookmarks(bookmark).Select()
        wordapp.Selection.TypeText(text)
    else:
            wordsel.InsertAfter(text)


def getStyleList(doc):
    # returns a dictionary of the styles in a document
    global styles
    stylecount = doc.Styles.Count
    for i in range(1, stylecount + 1):
        styleObject = doc.Styles(i)
        styles.append(styleObject.NameLocal)
    return styles

def addStyledPara(rnge, text, stylename):
##    if text[-1] <> '\n':
##        text = text + '\n'
    rnge.InsertAfter(text)
    rnge.Style = stylename
##    selectEnd()


wordapp = win32com.client.Dispatch("Word.Application")
wordapp.Visible = 1
wordsel = wordapp.Selection

worddoc = wordapp.Documents.Add()
worddoc.PageSetup.Orientation = 1
worddoc.PageSetup.BookFoldPrinting = 1
worddoc.Content.Font.Size = 11
worddoc.Content.Paragraphs.TabStops.Add (100)
worddoc.Content.Text = "Hello, I am a text!"

styles = []
textes = ['youpi1', 'youplaboum2']


for txt in textes:
    loc = worddoc.Range()
    loc.Paragraphs.Add()
    loc.Collapse(0)
    loc.InsertAfter(txt)
    loc.AutoFormat()


##location = worddoc.Range()
##location.Collapse(1)
##location.Paragraphs.Add()
##location.Collapse(1)
##table = location.Tables.Add (location, 3, 4)
##table.ApplyStyleHeadingRows = 1
##table.AutoFormat(16)
##table.Cell(1,1).Range.InsertAfter("Teacher")
##
##location1 = worddoc.Range()
##location1.Paragraphs.Add()
##location1.Collapse(1)
##table = location1.Tables.Add (location1, 3, 4)
##table.ApplyStyleHeadingRows = 1
##table.AutoFormat(16)
##table.Cell(1,1).Range.InsertAfter("Teacher1")
##worddoc.Content.MoveEnd


# styles
sty = worddoc.Styles
h1 = sty.Item(2)
getStyleList(worddoc)
print styles

addStyledPara(loc, '', u'Titre 1')


# table of contents
start = worddoc.Range(0,0)
start.Paragraphs.Add()
start.Collapse()
toc = worddoc.TablesOfContents
toc.Add(start)
toc1 = toc.Item(1)
toc1.IncludePageNumbers = True
toc1.RightAlignPageNumbers = True
toc1.UseFields = True
toc1.UseHeadingStyles = True

# Page numbers
sec = worddoc.Sections.Item(1)
bdp = sec.Footers.Item(1)
bdp.PageNumbers.Add()



### clean up
##worddoc.Close() # Close the Word Document (a save-Dialog pops up)
##wordapp.Quit() # Close the Word Application