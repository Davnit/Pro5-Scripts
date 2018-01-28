import xml.etree.ElementTree

# Returns the XML tree for the specified file (.PRO5)
def getDocumentTree(path):
    return xml.etree.ElementTree.parse(path)
    
# Returns a list of slide elements from a document tree root
def getSlides(document):
    return document.findall("./groups/RVSlideGrouping/slides/RVDisplaySlide")

# Returns a list of media cues contained on a slide
def getCues(slide):
    return slide.findall("./cues/RVMediaCue")
    