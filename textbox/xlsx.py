try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML
import zipfile


class XLSXExtractor(object):
    NAMESPACE = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    TEXT = NAMESPACE + 't'

    def __init__(self, filename):
        self.text = self._get_text(filename)                 

    def get_text(self):
        return self.text 

    def _get_text(self, filename):
        document = zipfile.ZipFile(filename)
        xml_content = document.read('xl/sharedStrings.xml')
        document.close()
        tree = XML(xml_content)

        texts = [node.text
                 for node in tree.getiterator(self.TEXT)
                 if node.text]

        return '\n\n'.join(texts)


def get_text(filename):
    xlsx = XLSXExtractor(filename)
    print(xlsx.get_text())


if __name__ == '__main__':
    import sys
    get_text(sys.argv[1])
