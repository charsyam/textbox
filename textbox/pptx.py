try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML
import zipfile


class PPTXExtractor(object):
    NAMESPACE = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    TEXT = NAMESPACE + 't'
    
    SLIDE_PREFIX = "ppt/slides/slide"
    SLIDE_PREFIX_LENGTH = len(SLIDE_PREFIX)
    
    def __init__(self, filename):
        self.text = self._get_text(filename)                 

    def get_text(self):
        return self.text 

    def _get_text(self, filename):
        document = zipfile.ZipFile(filename)
        texts = [self._get_text_from_slide(document, slide)\
                for slide in self._get_slide_names(document.namelist())]
        document.close()
        return '\n'.join(texts)

    def _get_slide_names(self, dirs):
        m = []
        for d in dirs:
            if d.startswith(self.SLIDE_PREFIX):
                m.append(int(d[self.SLIDE_PREFIX_LENGTH:-4]))

        return [self.SLIDE_PREFIX+str(x)+".xml" for x in sorted(m)]

    def _get_text_from_slide(self, document, slide):
        xml_content = document.read(slide)
        tree = XML(xml_content)

        texts = [node.text
                    for node in tree.getiterator(self.TEXT)
                    if node.text]

        return '\n'.join(texts)


def get_text(filename):
    pptx = PPTXExtractor(filename)
    print(pptx.get_text())


if __name__ == '__main__':
    import sys
    get_text(sys.argv[1])
