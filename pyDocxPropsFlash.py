#
# The PAIS Thesis Template
# Copyright (C) 2024 Roman Lupashko <mossy0.civets@icloud.com>
# 
# This file is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
# 
# This file is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
# 
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
# 
# GitHub repository <https://github.com/CuberHuber/thesis-word-template>
import argparse
import json

from docx import Document
from docx.opc.customprops import CustomProperties


class WordPropsFlash:
    def __init__(self, file: str, custom_props: str, new: str = None, force: bool = False):
        self.doc = None
        self.props = None
        self._run(file, custom_props, new, force)

    def properties(self, filepath: str):
        with open(filepath, 'r') as file:
            self.props = json.loads(file.read())

    def update_properties(self, document: Document, force: bool = False):
        crops: CustomProperties = document.custom_properties
        for prop in self.props.items():
            key = prop[0]
            if crops[key] is None and not force:
                print(f'key: {key} not in properties in current document')
            else:
                crops[prop[0]] = prop[1]

    def save_document(self, filepath: str):
        self.doc.save(filepath)

    def _run(self, file: str, custom_props: str, new: str = None, force: bool = False):
        print('Start updating')
        self.doc = Document(str(file))
        self.properties(custom_props)

        print('Properties loaded')
        self.update_properties(self.doc, force)

        self.save_document(new if new else file)
        print('Updating complete')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.description = '''Word document updator
    \r\n\t
    - это микроскрипт, который загружает в Word документ кастомные свойства из JSON файла'''
    parser.add_argument('--file', type=str, help='The absolute path to Word file')
    parser.add_argument('--custom_props', type=str, help='The absolute path to properties.json file')
    parser.add_argument('--new', type=str, help='The absolute path to new Word file')
    parser.add_argument('--force', action='store_true', help='Flag to force append new properties to document')
    args = parser.parse_args()

    word_doc = WordPropsFlash(args.file, args.custom_props, args.new, args.force)
