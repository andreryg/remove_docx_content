from lxml import etree
import zipfile
import os

def subparagraphToDocx(path_to_doc_xml_file, paragraph_header, paragraph_header_style, subparagraph_header=None, subparagraph_header_style=None):
    """
    This function removes all paragraphs from a Word document except for the one with the specified header.
    Optionally can also remove all subparagraphs except for the one with the specific subparagraph header.
    The subparagraph and paragraph headers must be in a specified style. This is probably a localized name.
    """

    NAMESPACE_PREFIXES = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    }

    tree = etree.parse(path_to_doc_xml_file)
    root = tree.getroot()

    text_elements = [element for element in root[0]]
    keep_element = False
    correct_paragraph = False
    for element in text_elements:    
        for subelement in element:
            if subelement.tag == '{' + NAMESPACE_PREFIXES['w'] + '}pPr':
                for subsubelement in subelement:
                    if subsubelement.tag == '{' + NAMESPACE_PREFIXES['w'] + '}pStyle':
                        if subsubelement.get('{' + NAMESPACE_PREFIXES['w'] + '}val') == paragraph_header_style and ''.join(element.itertext()).strip() == paragraph_header:
                            correct_paragraph = True
                            if not subparagraph_header:
                                root[0].remove(element)
                                keep_element = True
                        elif subparagraph_header and correct_paragraph and subsubelement.get('{' + NAMESPACE_PREFIXES['w'] + '}val') == subparagraph_header_style and ''.join(element.itertext()).strip() == subparagraph_header:
                            root[0].remove(element)
                            keep_element = True
                        elif subparagraph_header and subsubelement.get('{' + NAMESPACE_PREFIXES['w'] + '}val') == subparagraph_header_style:
                            keep_element = False
                        elif not subparagraph_header and subsubelement.get('{' + NAMESPACE_PREFIXES['w'] + '}val') == paragraph_header_style:
                            keep_element = False
        if not keep_element:
            root[0].remove(element)

    #Remove last pagebreak element
    if len(root[0]) > 0 and len(root[0][-1]) > 0 and len(root[0][-1][-1]) > 0 and len(root[0][-1][-1]) > 0:
        if root[0][-1][-1].tag == '{' + NAMESPACE_PREFIXES['w'] + '}br':
            root[0].remove(root[0][-1])

    with open(path_to_doc_xml_file, 'wb') as file:
        file.write(etree.tostring(tree, pretty_print=True, xml_declaration=True, encoding='UTF-8'))

def findSubsubparagraphs(path_to_doc_xml_file, subparagraph_header, subsubparagraph_header_style):
    """
    This function finds all subsubparagraph headers from a specific subparagraph and returns them as a list.
    The subparagraph header must be in a specified style. This is probably a localized name.
    """

    NAMESPACE_PREFIXES = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    }

    tree = etree.parse(path_to_doc_xml_file)
    root = tree.getroot()

    # Find all subsubparagraphs
    correct_subparagraph = False
    subsubparagraph_headers = []
    text_elements = [element for element in root[0]]
    for element in text_elements:
        if ''.join(element.itertext()).strip() == subparagraph_header:
            correct_subparagraph = True
        if correct_subparagraph and element.tag == '{' + NAMESPACE_PREFIXES['w'] + '}p':
            for subelement in element:
                if subelement.tag == '{' + NAMESPACE_PREFIXES['w'] + '}pPr':
                    for subsubelement in subelement:
                        if subsubelement.tag == '{' + NAMESPACE_PREFIXES['w'] + '}pStyle':
                            if subsubelement.get('{' + NAMESPACE_PREFIXES['w'] + '}val') == subsubparagraph_header_style:
                                subsubparagraph_header = ''.join(element.itertext()).strip()
                                if len(subsubparagraph_header) > 0:
                                    subsubparagraph_headers.append(subsubparagraph_header)
    return subsubparagraph_headers

def subsubparagraphToDocx(path_to_doc_xml_file, subsubparagraph_header, paragraph_header_style, subsubparagraph_header_style):
    """
    This function takes a subsubparagraph from a specific subparagraph in a specific paragraph and saves the subsubparagraph in a separate Word document.
    The paragraph needs to be in a table on contents. The subsubparagraph and paragraph headers must be in a specified style. This is probably a localized name.
    """

    NAMESPACE_PREFIXES = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    }

    tree = etree.parse(path_to_doc_xml_file)
    root = tree.getroot()

    #Create a new Word document for each subsubparagraph
    print(subsubparagraph_header)
    tree = etree.parse(path_to_doc_xml_file)
    root = tree.getroot()
    correct_subsubparagraph = False
    text_elements = [element for element in root[0]]
    for element in text_elements:
        for subelement in element:
            if subelement.tag == '{' + NAMESPACE_PREFIXES['w'] + '}pPr':
                for subsubelement in subelement:
                    if subsubelement.tag == '{' + NAMESPACE_PREFIXES['w'] + '}pStyle':
                        if subsubelement.get('{' + NAMESPACE_PREFIXES['w'] + '}val') == subsubparagraph_header_style and ''.join(element.itertext()).strip() == subsubparagraph_header:
                            root[0].remove(element)
                            correct_subsubparagraph = True
                        elif subsubelement.get('{' + NAMESPACE_PREFIXES['w'] + '}val') in [subsubparagraph_header_style, paragraph_header_style]:
                            correct_subsubparagraph = False
        #print(correct_subsubparagraph)#, ''.join(element.itertext()).strip())
        if not correct_subsubparagraph:
            root[0].remove(element)

    #Remove last pagebreak element
    if len(root[0]) > 0 and len(root[0][-1]) > 0 and len(root[0][-1][-1]) > 0 and len(root[0][-1][-1]) > 0:
        if root[0][-1][-1].tag == '{' + NAMESPACE_PREFIXES['w'] + '}br':
            root[0].remove(root[0][-1])
    
    #Save the new Word document
    with open(path_to_doc_xml_file, 'wb') as file:
        file.write(etree.tostring(tree, pretty_print=True, xml_declaration=True, encoding='UTF-8'))