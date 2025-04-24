from lxml import etree

def fixLinks(path_to_doc_xml_file):
    """
    This function takes a Word document and fixes the within-document links in the document.
    The reason why links needs to be fixed is that they are no longer working after extracting one paragraph to a new word document.
    The link texts are changed to format $linkText$ and the link is removed.
    """
    NAMESPACE_PREFIXES = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    }

    tree = etree.parse(path_to_doc_xml_file)
    root = tree.getroot()

    link_element = False
    text_elements = [element for element in root[0]]
    # Find all hyperlinks
    for element in text_elements:
        #All the links are in a table
        if element.tag == '{' + NAMESPACE_PREFIXES['w'] + '}tbl':
            for row in element:
                #Iterate table rows
                if row.tag == '{' + NAMESPACE_PREFIXES['w'] + '}tr':
                    for cell in row:
                        #Iterate table cells
                        if cell.tag == '{' + NAMESPACE_PREFIXES['w'] + '}tc':
                            for paragraph in cell:
                                #Iterate cell paragraphs
                                if paragraph.tag == '{' + NAMESPACE_PREFIXES['w'] + '}p':
                                    for run in paragraph:
                                        #Iterate paragraph runs
                                        if run.tag == '{' + NAMESPACE_PREFIXES['w'] + '}r':
                                            if run.find('{' + NAMESPACE_PREFIXES['w'] + '}fldChar') is not None:
                                                #Cleanup
                                                paragraph.remove(run)
                                            run_text = ''.join(run.itertext()).strip()
                                            if run_text[0:8] == "REF _Ref":
                                                paragraph.remove(run)
                                                link_element = True
                                            elif run_text and link_element:
                                                for text in run:
                                                    if text.tag == '{' + NAMESPACE_PREFIXES['w'] + '}t':
                                                        #Change link text to $linkText$
                                                        text.text = '$'+run_text+'$'
                                                        link_element = False
    
    with open(path_to_doc_xml_file, 'wb') as file:
        file.write(etree.tostring(tree, pretty_print=True, xml_declaration=True, encoding='UTF-8'))