try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML
import zipfile


class DOCXExtractor(object):
    NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    PARA = NAMESPACE + 'p'
    TEXT = NAMESPACE + 't'

    def __init__(self, filename):
        self.text = self._get_text(filename)                 

    def get_text(self):
        return self.text 

    def _get_text(self, filename):
        document = zipfile.ZipFile(filename)
        xml_content = document.read('word/document.xml')
        document.close()
        tree = XML(xml_content)

        paragraphs = []
        for paragraph in tree.getiterator(self.PARA):
            texts = [node.text
                    for node in paragraph.getiterator(self.TEXT)
                    if node.text]
            if texts:
                paragraphs.append(''.join(texts))

        return '\n\n'.join(paragraphs)


def get_text(filename):
    docx = DOCXExtractor(filename)
    print(docx.get_text())


if __name__ == '__main__':
    import sys
    get_text(sys.argv[1])
