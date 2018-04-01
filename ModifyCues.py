import sys
from ParsePRO5 import *

if len(sys.argv) < 5:
    sys.exit("You must specify a file path, xpath, target property, and new value.")
    
target = sys.argv[1]    # The path to the document file
xpath = sys.argv[2]     # The path to the node (under RVMediaCue) to be modified.
property = sys.argv[3]  # The name of the property to modify.
newValue = sys.argv[4]  # The new value to give the property.

# Get document data
tree = getDocumentTree(target)

# Do this for every cue on every slide
for slide in getSlides(tree.getroot()):
    for cue in getCues(slide):
        
        # Do it for every element matching the given path
        for element in cue.findall(xpath):
            # Change the value
            element.set(property, newValue)
            
# Save our changes
tree.write(target)