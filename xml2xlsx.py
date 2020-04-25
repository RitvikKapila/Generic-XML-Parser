#!/usr/bin/env python
# coding: utf-8


"""
Authors: Ritvik Kapila, Gauri Gupta
"""

import pandas as pd 
import numpy as np
import xml.etree.ElementTree as etree
import sys

# Counting the number of children of a given node

def total_entries(node):
    count = 0
    for i in range(sys.maxsize):
        try:
            node[i]
            count += 1
        except:
            return count 
        
# Removing unsupported characters from the sheet names in excel
def column_name(s):
    s1 = ""
    for i in range(len(s)):
        if not(s[i] == ']' or s[i] == '[' or s[i] == ':' or s[i] == '*' or s[i] == '?' or s[i] == '/'):
            s1 = s1 + s[i]
        else:
            s1 = s1 + "_"
    return s1 

# Renaming excel sheet if the name exceeds given characters 

def excel_name(d):
    count = 1
    d1 = {}
    for i in d:
        if len(i)>29:
            i1 = i[:29] + str(count)
            count = count + 1
            d1[i1] = d[i]
        else :
            d1[i] = d[i]
    return d1

# This function combines the data in child nodes of the list attributes and outputs a list

def entry_list(node):
    l1 = []
    for i in range(total_entries(node)):
        if node[i].text == "\n          ":
            ltemp = []
            for j in range(total_entries(node[i])):
                ltemp = ltemp + [node[i][j].text]
            l1.append(ltemp)    
        else :
            l1.append(node[i].text)
    return l1

# Converts the entire xml document to a dictionary containing different dataframes for each sheet attribute with keys as sheet names

def data_xml(root, sheet_attrib):
    
#     Creating the dictionary format
    keyList = []
    n = total_entries(root[0])
    for i in range(n):
        if i>0:
#             print(root[0][i].tag, "\n", root[0][i].attrib, "\n", root[0][i].text)
            keyList = keyList + [column_name(root[0][i].attrib[sheet_attrib])]
    dictionary = {}
    for i in keyList: 
        dictionary[i] = 'None'
#     print(dictionary)    
#     Adding data to the dictionary
    for i in range(n):
        if i>0:
            currNode = root[0][i]
            try:    
                if dictionary[column_name(currNode.attrib[sheet_attrib])] == 'None':
#                     print('empty' + currNode.attrib[sheet_attrib])
                    columns = []
                    values = []
                    for key in currNode.attrib:
                        if key == sheet_attrib:
                            continue
                        else:
                            columns = columns + [key]
                            values = values + [currNode.attrib[key]]
                    for j in range(total_entries(root[0][i])):
#                         print('yes')
                        if(currNode[j].tag == 'p' or currNode[j].tag == '{raml21.xsd}p'):
                            columns = columns + [currNode[j].attrib['name']]
                            values = values + [currNode[j].text]
                        elif (currNode[j].tag == 'list' or currNode[j].tag == '{raml21.xsd}list'):
                            columns = columns + [currNode[j].attrib['name']]
                            values = values + [entry_list(currNode[j])]

                    df = pd.DataFrame([values], columns = columns)
                    dictionary[column_name(currNode.attrib[sheet_attrib])] = df
#                     print(total_entries(currNode)) 
#                     print((dictionary))
            except:
#                 print('filled' + root[0][i].attrib[sheet_attrib])
                columns = []
                values = []
                for key in currNode.attrib:
                    if key == sheet_attrib:
                        continue
                    else:
                        columns = columns + [key]
                        values = values + [currNode.attrib[key]]                    
                for j in range(total_entries(root[0][i])):
                    if(currNode[j].tag == 'p' or currNode[j].tag == '{raml21.xsd}p'):
                        columns = columns + [currNode[j].attrib['name']]
                        values = values + [currNode[j].text]
                    elif (currNode[j].tag == 'list' or currNode[j].tag == '{raml21.xsd}list'):
                        columns = columns + [currNode[j].attrib['name']]
                        values = values + [entry_list(currNode[j])]
                                               
                df = pd.DataFrame([values], columns = columns)
                dictionary[column_name(currNode.attrib[sheet_attrib])] = dictionary[column_name(currNode.attrib[sheet_attrib])].append(df)
#                 print(total_entries(currNode)) 
#                 print((dictionary))
    return excel_name(dictionary)

# Converts an input xml file to output xlsx file given the sheet attribute for creating sheets in xlsx

def xml_to_xlsx(input_file, sheet_attrib, output_file):
    # Reading the xml as a tree object
    
    tree = etree.parse(input_file)
    root = tree.getroot()
    
    dfs = data_xml(root, sheet_attrib)
    
    # Opens the output file for writing

    writer = pd.ExcelWriter(output_file + '.xlsx', engine = 'xlsxwriter')
    
    # Writing the data from the dictionary onto the excel sheet

    for sheet_name in dfs.keys():
        dfs[sheet_name].to_excel(writer, sheet_name = sheet_name, index = False)

    writer.save()
    
# xml_to_xlsx('test3_xml.xml', 'class', 'out_3')    

