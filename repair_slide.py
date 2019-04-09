#! /usr/bin/env python
import os, sys
from lxml import etree
from shutil import rmtree
from opcdiag.controller import OpcController

def strQ2B(s):
    n = []
    s = s.decode('utf-8')
    for char in s:
        num = ord(char)
        if num == 0x3000:
            num = 32
        elif 0xFF01 <= num <= 0xFF5E:
            num -= 0xfee0
        num = unichr(num)
        n.append(num)
    return ''.join(n)

def repair_slide(target_path):
    # Step 1. Extract target_path
    opc = OpcController()
    TEMP_FOLDER_TARGET = target_path.replace(".pptx", '')
    opc.extract_package(target_path, TEMP_FOLDER_TARGET)

    vml_drawing = "{}/ppt/drawings/{}"
    xml_slide = "{}/ppt/slides/slide{}.xml"
    xml_slide_rel = "{}/ppt/slides/_rels/{}"

    # repair wmf position
    slide_rel_list = [x for x in os.listdir("{}/ppt/slides/_rels/".format(TEMP_FOLDER_TARGET))
                   if '.xml.rels' in x]
    for index1 in range(len(slide_rel_list)):
        with open(xml_slide_rel.format(TEMP_FOLDER_TARGET, slide_rel_list[index1])) as file:
            lines = [line for line in file.readlines() if '/drawings/vmlDrawing' in line]

        if len(lines) > 0:
            for line in lines:
                drawing_name = line.split('Target="../drawings/')[-1].replace('"/>\n', '')
                drawing_tree = etree.parse(vml_drawing.format(TEMP_FOLDER_TARGET, drawing_name))
                drawing_root = drawing_tree.getroot()
                shapeLst = drawing_root.findall("*[@type='#_x0000_t75']")

                slide_id = int(slide_rel_list[index1].replace("slide", "").replace(".xml.rels", ""))
                slide_tree = etree.parse(xml_slide.format(TEMP_FOLDER_TARGET, slide_id))
                slide_root = slide_tree.getroot()
                graphicFrameLst = slide_root.findall(".//p:graphicFrame",{'p': "http://schemas.openxmlformats.org/presentationml/2006/main"})
                if len(shapeLst) > 0:
                    for index2 in range(len(shapeLst)):
                        # print("before: " + shapeLst[index2].attrib["style"])
                        styleLst = shapeLst[index2].attrib["style"].split(";")
                        for index3 in range(len(styleLst)):
                            if "top:" in styleLst[index3]:
                                # print(styleLst[index3])
                                top = int(styleLst[index3].split(":")[1].replace("pt", ""))
                                styleLst[index3] = "top:" + str(top + 7) + "pt"
                                # print(styleLst[index3])
                        # print("styleLst: ")
                        # print(styleLst)
                        # increase top in style
                        shapeLst[index2].attrib["style"] = ';'.join(styleLst) + ";"
                        # print("after: " + shapeLst[index2].attrib["style"])
                        # increase y value in match graphicFrame's <a:off>
                        offLst = graphicFrameLst[index2].findall(".//a:off",{'a': "http://schemas.openxmlformats.org/drawingml/2006/main"})
                        for off in offLst:
                            off.attrib["y"] = unicode(int(off.attrib["y"]) * (top + 7) / top) 

                    # save drawing
                    with open(vml_drawing.format(TEMP_FOLDER_TARGET, drawing_name), 'w') as file:
                        file.writelines(etree.tostring(drawing_root))
                    # save slide
                    with open(xml_slide.format(TEMP_FOLDER_TARGET, slide_id), 'w') as file:
                        file.writelines(
                            "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>{}".format(
                                etree.tostring(slide_root)
                            )
                        )

    # repair string
    slide_list = [x for x in os.listdir("{}/ppt/slides/".format(TEMP_FOLDER_TARGET))
                   if '.xml' in x]
    for index1 in range(len(slide_list)):
        slide_tree = etree.parse("{}/ppt/slides/{}".format(TEMP_FOLDER_TARGET, slide_list[index1]))
        slide_root = slide_tree.getroot()
        tLst = slide_root.findall(".//a:t",{'a': "http://schemas.openxmlformats.org/drawingml/2006/main"})
        for t in tLst:
            text = t.text
            if isinstance (text, unicode): #unicode to string
                t.text = strQ2B(t.text.encode('utf-8'))
            else: 
                t.text = strQ2B(text)

        # save slide
        with open("{}/ppt/slides/{}".format(TEMP_FOLDER_TARGET, slide_list[index1]), 'w') as file:
            file.writelines(
                "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>{}".format(
                    etree.tostring(slide_root)
                )
            )

    opc.repackage(TEMP_FOLDER_TARGET, target_path.replace(".pptx", "-repair.pptx"))
    rmtree(TEMP_FOLDER_TARGET)

target_path = sys.argv[1]

repair_slide(target_path)
print(target_path.replace(".pptx", "-repair.pptx"))