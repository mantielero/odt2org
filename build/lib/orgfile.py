# -*- coding: utf-8 -*-
#!/usr/python
"""
Project Name: odt2org
Project hosted at http://bitbucket.org/josemaria.alkala/odt2org

Copyright 2010-2010 José María García Pérez

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

     http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
"""
import codecs
import os

class ORGfile:
    '''This class is used to generate an ORG file'''
    def __init__(self,
                 filename = None,
                 original = None,
                 overwrite = False):
        self.filename = os.path.realpath( filename )
        self.prefix = os.path.splitext(filename)[0]
        self.original = os.path.split(original)[1]
        self.overwrite = overwrite
        self.isWarning = False
        self.data = u'-*- mode: org; coding: utf-8 -*-\n'
        self.nrow = -1
    
    def read_list(self, _list, _extra):
        """
        """
        # Remove unprocessed stuff
        _newlist = []
        for _i in _list:
            if _i.__contains__('text'):
                _newlist.append(_i)
        _list = _newlist
        # First item should be the original filename     
        self.addHeading( level = 1,
                         text = u'%s' % self.original )

        # Are there "Title" styles?
        _hasTitle = False
        _n = 0
        for _i in _list:
            if _i['tags']['style'] == u'Title':
                _hasTitle = True
                _n = 1
        
        # Process
        for _i in _list:
            # Closing tables
            if _i['tags']['type'] != u'table-cell' and self.nrow != -1:
                self.data += '|\n'
                self.nrow = -1 
            # Title
            if _i['tags']['style'] == u'Title':
                self.addHeading( level = 2,
                                 text =  _i['text'])
            elif _i['tags']['type'] == u'heading':
                # TODO: 'level' should already arrive as Integer
                self.addHeading( level = int(_i['tags']['level'])+1+_n,
                                 text =  _i['text'])
            elif _i['tags']['type'] == u'paragraph':
                self.addParagraph( text = _i['text'] )
            elif _i['tags']['type'] == u'table-cell':
                self.add_cell( _i )
            elif _i['tags']['type'] == u'list-item':
                self.addEnumeration( list_item = _i )
            elif _i['tags']['type'] == u'image':
                self.add_image( _i)

        # Footnotes
        if _extra.__contains__('footnotes'):
            if _extra['footnotes'] != []:
                self.data += u'** Footnotes\n'
                for _i in _extra['footnotes']:
                    self.data += _i[0]
                    for _j in _i[1]:
                        self.data += _j['text'] +u'\n'
                    
    def add_image(self, _item):
        _ini = _item['text']
        _tmp = os.path.split(_ini)[1]
        _newfile = self.prefix+'_'+_tmp
        # TODO: check for overwrite
        if os.path.isfile( _newfile) and self.overwrite:
            os.remove(_newfile)
        else:
            raise OrgfileError(u"The file already exists: %s" % _newfile)
            
        os.rename(_ini,_newfile)

        _outputdir = os.path.split(self.filename)[0]
        _relpath = os.path.relpath( _newfile, _outputdir)
        _tmp = u'[[file:%s][%s]]\n' % (_relpath,_item['tags']['name'])
        self.data += _tmp

    def addHeading(self, 
                   level = 1,
                   text = u'' ):
        """Headings can only have one line in org-mode
        """
        text = text.replace('\r\n',' _ ')    
        text = text.replace('\n',' _ ')

        text = u'*'*(level) + u' %s\n' % text
        self.data = self.data + text

    def addParagraph(self, 
                     text = [] ):
        self.data = self.data + text + '\n'

    def add_cell(self, _cell):
        """
        """
        #print _cell
        _nrow = _cell['tags']['nrow']
        if _nrow == 0:
            self.nrow = _nrow

        if _nrow > self.nrow:
            self.data += u'|\n| %s ' % _cell['text']
            self.nrow = _nrow
        else:
            self.data += u'| %s ' %  _cell['text']

    def addEnumeration(self, 
                       list_item = {}):
        _level = int(list_item['tags']['level'])
        _tmp = u'  '*(_level -1) + u'- '+ list_item['text'] +'\n'
        self.data = self.data +  _tmp

    def export(self):
        _fp = codecs.open( self.filename, 'w', 'utf-8')
        _fp.write( self.data )
        _fp.close()

if __name__ == '__main__':
    pass
