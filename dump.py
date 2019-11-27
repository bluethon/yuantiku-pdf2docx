import argparse
import functools
import unicodedata
import re
from pathlib import Path
from typing import List, Optional
from enum import Enum

from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import (LAParams, LTPage, LTTextBoxHorizontal,
                             LTTextLineHorizontal)
from pdfminer.pdfdevice import PDFDevice
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.pdfpage import PDFPage, PDFTextExtractionNotAllowed
from pdfminer.pdfparser import PDFParser

from re import search


class SubjectType(Enum):
    cloze = '完形填空'
    read = '阅读理解'
    read_75 = '阅读理解7选5'


def uni_nornal(func):
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        res = func(*args, **kwargs)
        return unicodedata.normalize('NFKD', res)
    return wrapper


class PDFLine:
    debug = False

    def __init__(self, index: int, line: LTTextLineHorizontal):
        self._obj = line
        self._text: str = line.get_text()
        self.index = index

    def __repr__(self):
        return f'\t--- {self.index} --- {type(self._obj).__name__} ---\n'

    @property
    def is_first_line(self):
        return self.text.startswith(u' '*4)

    @property
    def is_next_line(self):
        return re.search(r'^\s\S+', self.text)

    @property
    @uni_nornal
    def text(self):
        return self._text


class PDFBox:
    debug = False

    def __init__(self, box: LTTextBoxHorizontal, page: int, last_obj: LTTextBoxHorizontal):
        self._obj = box
        self._last = last_obj
        self.index: str = f'{page}-{box.index}'
        # self.text = box.get_text()
        self.lines = self.get_lines()
        self.pars = self.convert_lines_to_pars()

    def __repr__(self):
        return f'--- {self.index} --- {type(self._obj).__name__} ---\n'

    @property
    def is_multilines(self):
        return len(self.lines) > 1

    @property
    def is_new(self):
        return len(self.lines) > 1

    @property
    @uni_nornal
    def text(self):
        if self.debug is True:
            return self._obj.get_text()
        return ''.join(self.pars)

    def get_lines(self):
        _lines: List[PDFLine] = []
        for i, line in enumerate(self._obj):
            _lines.append(PDFLine(i, line))
        return _lines

    def convert_lines_to_pars(self):
        _pars: List[str] = []
        for line in self.lines:
            # if line.is_first_line:
                # _pars.append(line.text)
            if len(_pars) > 0 and (re.search(r'^\s\S', line.text) or not _pars[-1][-2:] == '.\n'):
                _pars[-1] = _pars[-1][:-1] + line.text
            else:
                _pars.append(line.text)
        return _pars


class Subject:
    type: SubjectType
    title: str
    text: str
    questions: List[str]
    option_a: List[str]
    option_b: List[str]
    option_c: List[str]
    option_d: List[str]

    def __init__(self):
        self.title = ''
        self.text = ''
        self.questions = []
        self.option_a = []
        self.option_b = []
        self.option_c = []
        self.option_d = []


def parse_file(file: Path):
    with open(file, 'rb') as fp:
        parser = PDFParser(fp)
        doc = PDFDocument(parser)
        laparams = LAParams()
        text_boxes = []     # 清理后box列表

        if not doc.is_extractable:
            raise PDFTextExtractionNotAllowed

        rsrcmgr = PDFResourceManager()
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)

        last_out = None
        for i, page in enumerate(PDFPage.create_pages(doc)):
            interpreter.process_page(page)
            layout = device.get_result()

            for out in layout:
                if isinstance(out, LTTextBoxHorizontal):
                    text_boxes.append(PDFBox(out, i, last_out))
                else:
                    pass
    print('parse end')
    return text_boxes


def concat_subjects(boxes: List[PDFBox]):
    _subs = []

    # for box in boxes:
    # if box.text


def write_file(boxes: List[PDFBox], src_file: Path, filename: Path = None, flags='w'):
    result = False
    if filename is None:
        filename = src_file.with_suffix('.log')

    print(filename)

    try:
        with open(filename, flags) as fp:
            for box in boxes:
                fp.writelines(str(box))
                for line in box.lines:
                    fp.writelines(str(line))
                    fp.writelines(line.text)
                # fp.writelines(box.text)
            result = True
    except IOError as identifier:
        print(identifier)
    return result


def main():
    # encoding = 'utf-8'

    parser = argparse.ArgumentParser()
    parser.add_argument(dest='file')
    parser.add_argument('-d', dest='debug',
                        action='store_true', help='debug on')

    args = parser.parse_args()

    print(args)
    PDFBox.debug = args.debug
    PDFLine.debug = args.debug

    _file = Path(args.file)
    res_text = parse_file(_file)

    write_file(res_text, _file)
    print(_file)


if __name__ == '__main__':
    main()
