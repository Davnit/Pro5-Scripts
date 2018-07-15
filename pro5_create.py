from datetime import datetime
from urllib.parse import urlparse, quote
from xml.etree import ElementTree
import uuid, os, sys


def get_uuid():
    return uuid.uuid4()

def create_image_slide(width, height, foreground_path):
    # Make sure the file path is absolute
    foreground_path = os.path.abspath(foreground_path)

    # Split file name into parts
    s = os.path.splitext(os.path.basename(foreground_path))

    file_name = s[0]      # Name of the file without extension, for slide label.
    file_type = s[1][1:]  # File extension, for determining type

    type_lookup = {
        "jpg": "JPEG image", "jpeg": "JPEG image",
        "png": "Portable Network Graphics image"
    }

    # Lookup the file type based on extension. If it's not supported, return None.
    if file_type.lower() in type_lookup:
        file_type = type_lookup[file_type.lower()]
    else:
        return None

    # Determine the absolute URL of the foreground file.
    foreground_path = os.path.abspath(foreground_path)
    path_uri = urlparse(foreground_path)
    if len(path_uri.netloc) == 0:
        path_uri = urlparse(
            "file://localhost/" + (path_uri.scheme + "/" if len(path_uri.scheme) == 1 else "") + quote(path_uri.path[1:].replace("\\",  "/")))

    slide = {
        "RVDisplaySlide": {
            "backgroundColor": "0 0 0 1", "enabled": "1", "highlightColor": "0 0 0 0", "hotKey": "", "label": file_name,
            "notes": "",
            "slideType": "1", "sort_index": "0", "UUID": get_uuid(), "drawingBackgroundColor": "0",
            "chordChartPath": "", "serialization-array-index": "0",

            "cues": {
                "RVMediaCue": {
                    "displayName": file_name, "delayTime": "0", "timeStamp": "0", "enabled": "1", "UUID": get_uuid(),
                    "parentUUID": "",
                    "elementClassName": "RVImageElement", "behavior": "2", "alignment": "4",
                    "serialization-array-index": "0",

                    "element": {
                        "displayDelay": "0", "displayName": file_name, "locked": "0", "persistent": "0", "typeID": "0",
                        "fromTemplate": "0",
                        "bezelRadius": "0", "drawingFill": "0", "drawingShadow": "0", "drawingStroke": "0",
                        "fillColor": "1 1 1 1",
                        "rotation": "0", "source": path_uri.geturl(), "flippedHorizontally": "0",
                        "flippedVertically": "0", "scaleFactor": "1",
                        "serializedImageOffset": "0.000000@0.000000", "serializedFilters": "", "scaleBehavior": "0",
                        "brightness": "0",
                        "contrast": "1", "saturation": "1", "hue": "0", "manufactureURL": "", "manufactureName": "",
                        "format": file_type,
                        "enableColorFilter": "0", "colorFilter": "1 0 0 1", "enableBlur": "0", "edgeBlurRadius": "0",
                        "edgeBlurArea": "0",
                        "enableSepia": "0", "enableColorInvert": "0", "enableGrayInvert": "0",
                        "enableHeatSignature": "0",

                        "_-RVRect3D-_position": { "x": "0", "y": "0", "z": "0", "width": width, "height": height },
                        "_-D-_serializedShadow": {
                            "containerClass": "NSMutableDictionary",
                            "NSMutableString": { "serialization-native-value": "{5, -5}",
                                                "serialization-dictionary-key": "shadowOffset" },
                            "NSNumber": { "serialization-native-value": "0",
                                         "serialization-dictionary-key": "shadowBlurRadius" },
                            "NSColor": { "serialization-native-value": "0 0 0 0.3333333432674408",
                                        "serialization-dictionary-key": "shadowColor" }
                        },
                        "stroke": {
                            "containerClass": "NSMutableDictionary",
                            "NSColor": { "serialization-native-value": "0 0 0 1",
                                        "serialization-dictionary-key": "RVShapeElementStrokeColorKey" },
                            "NSNumber": { "serialization-native-value": "1",
                                         "serialization-dictionary-key": "RVShapeElementStrokeWidthKey" }
                        }
                    },

                    "_-RVProTransitionObject-_transitionObject": {}
                }
            },

            "displayElements": { "containerClass": "NSMutableArray" },

            "_-RVProTransitionObject-_transitionObject": {
                "transitionType": "-1", "transitionDuration": "1", "motionEnabled": "0", "motionDuration": "20",
                "motionSpeed": "100"
            }
        }
    }

    slide["RVDisplaySlide"]["cues"]["RVMediaCue"]["parentUUID"] = slide["RVDisplaySlide"]["cues"]["RVMediaCue"]["UUID"]
    return dict_to_xml(slide)


def create_document(width, height, category):
    today = datetime.now().isoformat('T')

    return ElementTree.ElementTree(dict_to_xml({
        "RVPresentationDocument": {
            "height": height, "width": width, "versionNumber": "500", "docType": "0", "creatorCode": "1349676880",
            "lastDateUsed": today, "usedCount": "0", "category": category, "resourcesDirectory": "",
            "backgroundColor": "0 0 0 1",
            "DrawingBackgroundColor": "0", "notes": "", "artist": "", "author": "", "album": "", "CCLIDisplay": "0",
            "CCLIArtistCredits": "", "CCLIPublisher": "", "CCLISongTitle": "", "CCLICopyrightInfo": "",
            "CCLILicenseNumber": "",
            "chordChartPath": "",

            "timeline": {
                "timeOffset": "0", "selectedMediaTrackIndex": "0", "unitOfMeasure": "1", "duration": "0", "loop": "0",

                "timeCues": { "containerClass": "NSMutableArray" },
                "mediaTracks": { "containerClass": "NSMutableArray" }
            },

            "bibleReference": { "containerClass": "NSMutableDictionary" },

            "_-RVProTransitionObject-_transitionObject": {
                "transitionType": "-1", "transitionDuration": "1", "motionEnabled": "0", "motionDuration": "20",
                "motionSpeed": "100"
            },

            "groups": {
                "containerClass": "NSMutableArray",

                "RVSlideGrouping": {
                    "name": "", "uuid": get_uuid(), "color": "0 0 0 0", "serialization-array-index": "0",

                    "slides": { "containerClass": "NSMutableArray" }
                }
            },

            "arrangements": { "containerClass": "NSMutableArray" }
        }
    }))


def dict_to_xml(d, key=None):
    if key is None:
        if len(d.values()) > 1:
            raise Exception("Error: root dictionary must only have one value")
        else:
            key = next(iter(d.keys()))
            d = d[key]

    attr = {}
    subs = []
    for k, v in d.items():
        if type(v) == dict:
            subs.append(dict_to_xml(v, k))
        else:
            attr[k] = str(v)

    e = ElementTree.Element(key, attr)
    e.extend(subs)
    return e

def import_images_to_document(document, image_list):
    width = document.getroot().attrib["width"]
    height = document.getroot().attrib["height"]

    # Get slide elements for each image
    slides = [ create_image_slide(width, height, file) for file in image_list if os.path.isfile(file) ]

    # Add each file in the source path to the document's "slides" element
    e = document.find("./groups/RVSlideGrouping/slides")
    e.extend([ s for s in reversed(slides) if s is not None ])


# Main script
if len(sys.argv) > 1:
    doc_path = sys.argv[1]
    src_path = os.path.abspath(sys.argv[2])
    doc_width = sys.argv[3]
    doc_height = sys.argv[4]
    category = sys.argv[5]

    # Create a new XML base document
    doc = create_document(doc_width, doc_height, category)
    import_images_to_document(doc, os.listdir(src_path))

    # Write the new document
    doc.write(doc_path)