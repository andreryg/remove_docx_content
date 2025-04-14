from lxml import etree


def removeAllButOneParagraph(path_to_doc_xml_file, paragraph_header, subparagraph_header=None, subparagraph_header_style=None):
    """
    This function removes all paragraphs from a Word document except for the one with the specified header.
    Optionally can also remove all other subparagraphs except for the one with the specific subparagraph header.
    The subparagraph header must be in the same style as the specified subparagraph header style. This is probably a localized name.

    The logic behind choosing paragraphs is based on headers in the table of contents. 
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
            #The '_Toc' is a marker for the table of contents
            if subelement.tag == '{' + NAMESPACE_PREFIXES['w'] + '}bookmarkStart' and '_Toc' in subelement.get('{' + NAMESPACE_PREFIXES['w'] + '}name'):
                if ''.join(element.itertext()) == paragraph_header:
                    correct_paragraph = True
                    if subparagraph_header:
                        keep_element = False
                    else:
                        keep_element = True
                else:
                    keep_element = False
                    correct_paragraph = False
            if len(subelement) > 0 and correct_paragraph and subparagraph_header:
                for subsubelement in subelement:
                    if subsubelement.tag == '{' + NAMESPACE_PREFIXES['w'] + '}pStyle':
                        if subsubelement.get('{' + NAMESPACE_PREFIXES['w'] + '}val') == subparagraph_header_style:
                            if ''.join(element.itertext()).strip() == subparagraph_header:
                                keep_element = True
                            else:
                                keep_element = False
        if not keep_element:
            root[0].remove(element)

    #Remove last pagebreak element
    if root[0][-1][-1][-1].tag == '{' + NAMESPACE_PREFIXES['w'] + '}br':
        root[0].remove(root[0][-1])

    with open(path_to_doc_xml_file, 'wb') as file:
        file.write(etree.tostring(tree, pretty_print=True, xml_declaration=True, encoding='UTF-8'))