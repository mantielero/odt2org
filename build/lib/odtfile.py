# -*- coding: utf-8 -*-
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
ERROR = False
import codecs
import os
import string
import re
import tempfile
from orgfile import ORGfile

try:
    import zipfile
except ImportError:
    print u"ERROR: package 'zipfile' is mandatory. Install Python2.6 or higher"
    ERROR=True

try:
    from lxml import etree
except ImportError:
    print u'ERROR: install "nxml": http://codespeak.net/lxml/'
    ERROR = True

try:
    import OleFileIO_PL as ole
except ImportError:
    print u'ERROR: please, install "OleFileIO_PL": http://www.decalage.info/files/OleFileIO_PL-0.18.zip'
    # http://www.decalage.info/en/python/olefileio
    ERROR = True

class ODTfileError(Exception):
    def __init__(self):
        super(ODTfileError,self).__init__()

class ODTfile:
    '''This class enables the access to OpenOffice files'''
    def __init__(self,
                 filename = None):
        '''All the fields are mandatory. All the checks are done outside.
        '''
        self.tmpdir = tempfile.mkdtemp()
        self.ns = { 
           'office' : "urn:oasis:names:tc:opendocument:xmlns:office:1.0", 
           'text'   : "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
           'table'  : "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
           'draw'   : "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
           'xlink'  : "http://www.w3.org/1999/xlink"
                  }

        self.types = { 
            '{%s}%s' % (self.ns['text'],'p') : 'paragraph',
            '{%s}%s' % (self.ns['text'],'h') : 'heading',
            '{%s}%s' % (self.ns['text'],'outline-level') : 'outline-level',
            '{%s}%s' % (self.ns['text'],'list'): 'list',
            '{%s}%s' % (self.ns['text'],'list-item'): 'list-item',
            '{%s}%s' % (self.ns['text'],'style-name'): 'style-name',
            '{%s}%s' % (self.ns['text'],'s') : 'spaces',
            '{%s}%s' % (self.ns['text'],'tab') : 'tabs',
            '{%s}%s' % (self.ns['text'],'span') : 'span',
            '{%s}%s' % (self.ns['text'],'note') : 'note',
            '{%s}%s' % (self.ns['text'],'note-citation') : 'note-citation',
            '{%s}%s' % (self.ns['text'],'note-body') : 'note-body',
            '{%s}%s' % (self.ns['text'],'soft-page-break') : 'soft-page-break',
            '{%s}%s' % (self.ns['draw'],'frame') : 'frame',
            '{%s}%s' % (self.ns['draw'],'name') : 'framename',
            '{%s}%s' % (self.ns['draw'],'object-ole'): 'object-ole',
            '{%s}%s' % (self.ns['draw'],'image'): 'image',
            '{%s}%s' % (self.ns['xlink'],'href'): 'href',
            '{%s}%s' % (self.ns['table'],'table'): 'table',
            '{%s}%s' % (self.ns['table'],'table-row'): 'table-row',
            '{%s}%s' % (self.ns['table'],'table-cell'): 'table-cell'
                        }

        self.garbage = ['soft-page-break']

        if filename != None:
            self.filename = os.path.realpath( filename )

            # Prefix
            #_tmp = os.path.split( self.orgfile.filename )[1]
            #_tmp = _tmp.replace(' ','_') 
            #_tmp = _tmp.replace('.','_')
            #self.prefix = _tmp + '_'

            # Files
            try:
                self.files = zipfile.ZipFile( self.filename, "r")
            except:
                raise ODTFileError( u"""ODT files are zip files. The file provided isn't.""" )

            if not self.isOpenOfficeFile():
                raise ODTFileError( u"""Only compatible with ODT files 1.0 specification.""" )
 
            self.xml = None
        else:
            raise ODTFileError( u"""Provided ODT file class ODTfile requires a valid filename during initialization.""" )

    def isOpenOfficeFile(self):
        '''Verifies if this is an Open Office Document'''
        # 1. Contiene el fichero "content.xml"
        _isOK = False
        if self.files.namelist().__contains__( 'content.xml'):

            # 2. Leemos el fichero 'content.xml' y vemos que tiene el tag "office:document-content"
            _data = self.files.read( 'content.xml' )
            self.xml = etree.fromstring( _data )
            _lista = self.xml.xpath('/office:document-content', 
                                namespaces= { 'office' : 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'} )
            if len(_lista) == 1:
                _isOK = True
        return _isOK 

    def _get_textbody( self ):
        """Returns the body from the file
        """
        _data = self.files.read('content.xml')
        _dataxml = etree.fromstring( _data )

        _body = _dataxml.xpath('/office:document-content/office:body', 
                             namespaces= { 'office' : 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'} )[0]

        return _body.xpath('office:text', 
                            namespaces= { 'office' : 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'} )[0]


    def gen_list(self):
        """Creates a list that is understood by ORGfile class.
        """
        # Get all the nodes.
        _xml = self._get_textbody()
        _list = []

        for _child in _xml.iterdescendants():
            _level = 0
            for _i in _child.iterancestors():
                _level += 1
            
            _tmp = { 'child' : _child,
                     'pretext' : _child.text,
                     'posttext' : _child.tail,
                     'nesting' : _level}
            _list.append( _tmp )

        # Get warnings: information not managed
        _warnings = []
        for _i in _list:
            try:
                self.types[_i['child'].tag]
            except:
                _warnings.append(_i['child'].tag)
        _warnings = set(_warnings)
        print u"WARNING: following items are present in the document and won't be interpreted:"
        for _i in _warnings:
            print u"   %s" % _i
        print u"\n"


        # Ensure the list only contains info that we process.
        _newlist = []
        for _i in _list:
            try:
                _type = self.types[_i['child'].tag]
                if not self.garbage.__contains__( _type):
                    _newlist.append(_i)
            except:
                pass
        _list = _newlist

        # Extract text from paragraphs, heading, spaces, tabs, ...

        # - Extract spaces, tabs
        _newlist = []
        for _i in _list:
            if _i.__contains__('child'):
                _child = _i['child']
                _tag = _child.tag
                if self.types.__contains__(_tag):
                    _type = self.types[_tag]
                    if _type == 'spaces':
                        _txt = self.get_spaces(_child)
                        if _i['pretext'] != None:
                            _txt = _i['pretext'] + _txt
                        if _i['posttext'] != None:
                            _txt = _txt + _i['posttext']

                        _tmp = {'txt' : _txt, 'nesting': _i['nesting']}
                        _newlist.append(_tmp)
                    elif _type == 'tabs':
                        _txt = self.get_tabs(_child)
                        if _i['pretext'] != None:
                            _txt = _i['pretext'] + _txt
                        if _i['posttext'] != None:
                            _txt = _txt + _i['posttext']
                        _tmp = {'txt' : _txt, 'nesting': _i['nesting']}
                        _newlist.append(_tmp)
                    else:
                        _newlist.append(_i)
                else:
                    _newlist.append(_i)
            else:
                _newlist.append(_i)
            
        _list = _newlist  

        # - Extract text:span
        _newlist = []
        for _i in _list:
            try:
                if self.types[_i['child'].tag] == 'span':
                    _txt = u''
                    if _i['pretext'] != None:
                        _txt += _i['pretext']
                    if _i['posttext'] != None:
                        _txt += _i['posttext']
                    _tmp = {'txt' : _txt, 'nesting': _i['nesting']-1}
                    _newlist.append(_tmp)
                else:
                    _newlist.append(_i)
            except:
                _newlist.append(_i)

        _list = _newlist

        # - Footnotes
        _footnotes = []
        _list = self.group_type( _list, tag = 'note')
        _newlist = []
        for _i in _list:
            if _i.__contains__('note'):
                _txt, _paragraphs = self.get_footnote(_i)
                _newlist.append( {'txt' : _txt, 
                                  'nesting' : _i['nesting']})
                
                _footnotes.append((_txt, _paragraphs))
            else:
                _newlist.append(_i)
        _list = _newlist

        # - Regroup: move spaces, tabs into the prior child.
        _list = self.process_regroup(_list)
         
        # - Getting heading and paragraph text.
        _list = self.process_headings_paragraphs( _list )
        # LIST:
        _list = self.group_children( _list, tag = 'list', clean= True)
        _list = self.process_enumeration(_list)

        # TABLE
        _list = self.group_children( _list, tag = 'table', clean= True)
        _list = self.process_tables(_list)

        # FRAME
        # - Don't clean to recover the framename
        _list = self.group_children( _list, tag = 'frame', clean= False)
        _list = self.process_framename(_list)

        # Process Footnotes
        _tmp = []
        for _i in _footnotes:
            _tmp1 = self.process_regroup( _i[1] )

            _tmp1 = self.process_headings_paragraphs( _tmp1 )
            _tmp.append((_i[0],_tmp1))

        _extra = {'footnotes' : _tmp}
        return _list, _extra

    def process_regroup(self,_list):
        _newlist = []
        _tmp = u''
        for _i in xrange(len(_list)):
            _item = _list[_i]
            if _item.__contains__('child'):
                if _tmp != u'':
                    _newlist[-1].update({'txt' : _tmp})
                    _tmp = u''
                _newlist.append(_item)
            else:
                _tmp = _tmp + _item['txt']
            
        return  _newlist        

    def process_headings_paragraphs( self,_list ):
        _newlist = []
        for _i in _list:
            _isHeading = False
            _isParagraph = False

            try:
                if self.types[ _i['child'].tag] == 'heading':
                    _isHeading = True
            except:
                pass 
            try:
                if self.types[ _i['child'].tag] == 'paragraph':
                    _isParagraph = True
            except:
                pass 
            
            if _isHeading or _isParagraph:                    
                _tmp = u''
                if _i['pretext'] != None:
                    _tmp += _i['pretext']
                if _i.__contains__('txt'):
                    if _i['txt'] != None:
                        _tmp += _i['txt']
                if _i['posttext'] != None:
                    _tmp += _i['posttext']
                _tags = self.get_tags( _i['child'])
                if _isHeading:
                    _tags.update({'type' : 'heading'})
                elif _isParagraph:
                    _tags.update({'type' : 'paragraph'})

                _tmp = { 'text' : _tmp,
                         'nesting': _i['nesting'],
                         'tags' : _tags}

                _newlist.append(_tmp)

            else:
                _newlist.append(_i)
            
        return _newlist        

    def process_enumeration(self, _list):
        """Convert lists into something closer to text.
        Structure for Enumerations:
        <text:list>
           <text:list-item>
                <text:p>
        """
        # Phase 1
        _newlist = []
        for _i in _list:
            if _i.__contains__('list'):
                # - Clean "list"
                _tmp = self.clean( _i['list'], tags = ['list'])
                # - Group "list-item"
                _tmp = self.group_type( _tmp, tag = 'list-item', clean= True)
                _tmp1 = _i.copy()
                _tmp1['list'] = _tmp

                _newlist.append( _tmp1)
            else:
                _newlist.append( _i)
        _list = _newlist

        # Phase 2
        _newlist = []
        for _i in _list:
            if _i.__contains__('list'):
                _tmp = []

                for _j in _i['list']:
                    if _j.__contains__('list-item'):
                        _paragraphs = _j['list-item']
                        _txt = u''
                        for _k in _paragraphs:
                            _txt += _k['text'] +'\n'

                        if len(_paragraphs) > 0:
                            _txt = _txt[0:-1]
                            _tmp1 = _paragraphs[0].copy()
                            _tmp1['text'] = _txt

                            _tmp1['tags']['type'] = 'list-item'
                            _n = (_tmp1['nesting']-_i['nesting'] )/ 2
                            _tmp1['tags']['level'] = _n
                            _tmp.append( _tmp1)

                _newlist += _tmp
            else:
                _newlist.append( _i )

        return _newlist

    def process_tables(self, _list):
        """Simplies the format in which tables are expressed.
        """
        # Phase 1   
        _newlist = []
        for _i in _list:
            if _i.__contains__('table'):
                # - Group "table-row"
                _tmp = self.group_children( _i['table'], 
                                             tag = 'table-row', 
                                             clean= True)

                _tmp1 = []
                for _j in _tmp:
                    if _j.__contains__('table-row'):
                        _tmp2 = self.group_children( _j['table-row'], 
                                             tag = 'table-cell', 
                                             clean= True)
                        _tmp3 = _j.copy()
                        _tmp3['table-row'] = _tmp2
                        _tmp1.append( _tmp3)
                    else: 
                        _tmp1.append( _j)

                _tmp4 = _i.copy()
                _tmp4['table'] = _tmp1
                _newlist.append( _tmp4)
            else:
                _newlist.append( _i)
        _list = _newlist
        
        # Phase 2: convert into "text"
        _newlist = []
        for _i in _list:
            if _i.__contains__('table'):
                _table = _i['table']
                # Get rows.
                _nrow = 0
                for _j in _table:
                    if _j.__contains__('table-row'):
                        _row = _j['table-row']
                        # Get cells
                        for _k in _row:
                            if _k.__contains__('table-cell'):
                                _cell = _k['table-cell']
                                _txt = u''
                                for _m in _cell:
                                    _txt += _m['text'] + ' '
                                _txt = _txt[0:-1]
                                _tmp = _cell[0].copy()
                                _tmp['text'] = _txt
                                _tmp.pop('nesting')
                                _tmp['tags']['type'] = 'table-cell'
                                _tmp['tags'].update( { 'nrow' : _nrow })
                                _newlist.append( _tmp)
                        _nrow +=1
            else:
                _newlist.append( _i ) 
        return _newlist

    def process_framename(self, _list):
        # Phase 1
        _newlist = []
        _counter = 0
        for _i in _list:
            if _i.__contains__('frame'):
                _frame = _i['frame']
                # Framename
                _tags = self.get_tags(_frame[0]['child'])
                if not _tags.__contains__('framename'): 
                    _tags.update({'framename' : None}) 
                # - Images
                _tmp = self.group_children(_frame[1:],tag = 'image',clean = False)
                _tmp1 = []
                for _j in _tmp:
                    if _j.__contains__('image'):
                        _images = _j['image']
                        _n = 0
                        for _image in _images:
                            # - Extract reference
                            _imgtags = self.get_tags( _image['child'])
                            if _imgtags.__contains__('href'):
                                _dict = { 'text' : _imgtags['href'],
                                          'tags' : {'type' : 'image',
                                                    'style' : None} }
                                if _tags['framename'] == None:
                                    _name = 'Img_%.4d' % _counter
                                    _counter += 1
                                else:
                                    if len(_images) >1:
                                        _name = '%s_%.4d' % (_tags['framename'],_n)
                                        _n +=1
                                    else:
                                        _name = '%s' % _tags['framename']
                                _dict['tags'].update({'name': _name})
                                _tmp1.append(_dict)
                    else:
                        _tmp1.append(_j)
                _tmpp = _i.copy()
                _tmpp['frame'] = _tmp1 
                _newlist.append(_tmpp)
            else:
                _newlist.append( _i )
        _list = _newlist

        # Phase 2: create text
        _newlist = []
        for _i in _list:
            if _i.__contains__('frame'):
                _frame = _i['frame']
                for _j in _frame:
                    if _j.__contains__('text'):
                        _newlist.append(_j)
            else:
                _newlist.append(_i)
        _list = _newlist

        # Phase 3: file extractor
        _newlist = []
        for _i in _list:
            if _i.__contains__('tags'):
                if _i['tags']['type'] == 'image':
                    _newfilename = self.__fileExtractor__( _i['text'] )
                    _i['text'] = _newfilename
                    #print _newfilename
                    _newlist.append(_i)
                else:
                    _newlist.append(_i)
            else:
                _newlist.append(_i)

        # - Si el filename contiene: 'ObjectReplacements/' entonces no se extrae.
        # antes usaba: is_object_ole y también is_image, pero no entiendo con qué propósito.
        #_newfilename = self.__fileExtractor__( _filename )
        # '[[file:%s][%s]]\n' % (_newfilename,_name)

        return _newlist

    def clean(self, _list, tags = []):
        """Removes those items tagged with something contain in the list 'tags'
        """
        _newlist = []
        for _i in _list:
            try:
                _type = self.types[ _i['child'].tag ]
            except:
                _type = None
            if not tags.__contains__(_type):
                _newlist.append(_i)
        #if tags == ['list']:
        #    print _newlist,'\n\n'
        return _newlist     

    def get_istagged(self,_list,tag = 'list'):
        """Creates an array stating 'True' where the tag is as specified.
        """
        _tmp = []
        for _i in _list:
            try:
                if self.types[ _i['child'].tag ] == tag:
                    _tmp.append( True )
                else:
                    _tmp.append( False )
            except:
                _tmp.append( False )
        return _tmp

    def get_number_of_children(self, _list, idx = 0):
        """Given an item in a list indicated by its position: 'idx', this
        function returns the number of children it has.
        """
        _nesting = _list[idx]['nesting']
        _count = 0
        for _i in xrange(idx+1, len(_list)):
            try:
                _list[_i]['nesting']
            except:
                print u"ERROR: it should contains 'nesting' information:\n", _list[_i]
            if _list[_i]['nesting'] > _nesting:
                _count += 1
            else:
                break
        return _count

    def get_pairs(self,_list, tag= 'list'):  #<<<<<<<<<<<<<<<
        """Creates a list with pairs. Each pairs represent and item having
        an specified tag and ALL its children (even if the tag is contained
        again among its chuldren).
        """
        _istagged = self.get_istagged( _list,tag) 
        #if tag == 'table-row':
        #    print _istagged

        #print _istagged
        _pairs = []
        _n = 0
        for _i in xrange(len(_list)):
            if len(_pairs) > 0:
                _limit = _pairs[-1][1]
            else:
                _limit = 0
            #if tag == 'list':
            #    print _pairs
            #    print _i, _limit
            if _istagged[_i] and _i >= _limit:
                _nesting = _list[_i]['nesting']
                _n = self.get_number_of_children( _list, idx = _i)
                _pairs.append( [_i, _i+_n+1])
        return _pairs

    def group_children( self,
                        _list, 
                        tag = 'list',
                        clean = False ):
        """
        """
        if clean:
            _tmp = 1
        else:
            _tmp = 0
        _pairs = self.get_pairs(_list, tag)
        _newlist = []
        _ini = 0
        for _pair in _pairs:
            _newlist += _list[_ini:_pair[0]]
            _newlist.append( { tag : _list[_pair[0]+_tmp:_pair[1]], 
                               'nesting' : _list[_ini]['nesting']})
            _ini = _pair[1]
        _newlist += _list[_ini:]
        return _newlist

    def group_type( self,
                    _list, 
                    tag = 'note', 
                    clean = False ):
        _pairs = self.extract_children(_list, tag )
        if tag == 'list':
            print '    PAIRS: ',_pairs
        # Extract _list
        _newlist = []
        _oldidx = 0
        for _pair in _pairs:
            _newlist += _list[_oldidx:_pair[0]]
            _note = _list[_pair[0]:_pair[1]]
            _posttext = None
            if _note[0]['posttext'] != None:
                _posttext = _note[0]['posttext']
                _note[0]['posttext'] = None
            _nesting = _list[_pair[0]]['nesting']
            if clean:
                _note = _note[1:]
            _newlist.append( { tag : _note,
                               'nesting' : _nesting})
            if _posttext != None:
                _newlist.append( {'txt' : _posttext,
                                  'nesting' : _nesting})
            _oldidx = _pair[1]
        _newlist += _list[_oldidx:]
        return _newlist

    def extract_children( self, _list, tag = 'note'):
        _newlist = []
        _limits = []
        _flag = False
        _nesting = 0
        #if tag == 'list-item':
        #   print '== OK ==', _limits
        for _i in xrange(len(_list)):
            _item = _list[_i]
            try:
                if self.types[ _item['child'].tag ] == tag:
                    #_test = groupchildren and _item['nesting']< _nesting
                    #if not _test:
                    if _flag:
                        _limits.append(_i)
                        _newlist.append( _limits)
                    _flag = True
                    _nesting = _item['nesting']+1
                    _limits = [_i]                    
                    _tmp = _item
            except:
                pass
       
            if _flag and _item != _tmp:
                if _item['nesting'] < _nesting:
                    _flag = False
                    _nesting = 0
                    _limits.append(_i)

                    _newlist.append( _limits)
                    _limits = []
        if _limits != []:
            _limits.append(len(_list))
            _newlist.append( _limits)
         
        return  _newlist        

    def get_footnote(self, _note):
        _paragraphs = []
        _flag = False

        for _item in _note['note']:
            if _flag:
                _paragraphs.append(_item)
            try:
                _type = self.types[_item['child'].tag]
            except:
                _type = ''
            if _type == 'note-citation':
                _id = _item['pretext']
            elif _type == 'note-body':
                _flag = True
        _id = _id.replace('[','_')
        _id = _id.replace(']','_')
        _id = _id.replace(':','_')
        _id = u'[fn:%s] ' % _id
        #print _paragraphs
        return _id, _paragraphs

    def gen_list2(self):
        """Creates a list that is understood by ORGfile class.
        """
        _xml = self._get_textbody()
        _list = []

        for _child in _xml.getchildren():
            _tmp = self.get_newlist( [{'child' : _child,
                                       'tags' : {'nesting' : 1}}] )
            if _tmp != None:
                _list = _list + _tmp
        for _i in _list:
            print _i
        return _list


    def get_newlist(self, _list ):
        """
        """
        while self.has_children(_list):
            _newlist = []
            for _i in _list:
                if not _i.__contains__('child'):
                    _newlist.append(_i)
                else:
                    _newlist += self.analyse_child( _i)
            _list = _newlist
        return _list 

    def has_children(self, _list):
        for _i in _list:
            if _i.__contains__('child'):
                return True
        return False


    def get_tags(self, _child):
        """This function creates a tagged text
        """
        _dict = _child.attrib
        _tags = {}

        # Style
        _tagstyle = '{%s}%s' % (self.ns['text'],'style-name')
        if _dict.__contains__(_tagstyle):
            _tagname = _child.attrib[_tagstyle]
            _tags.update({'style' : _tagname } )
        else:
            _tags.update({'style' : None} )

        # Level
        _taglevel = '{%s}%s' % (self.ns['text'],'outline-level')
        if _dict.__contains__(_taglevel):
            _level = _child.attrib[_taglevel]
            _tags.update({'level' : _level })
        else:
            _tags.update({'level' : None} )

        # Framename
        _tag = '{%s}%s' % (self.ns['draw'],'name')
        if _dict.__contains__(_tag):
            _value = _child.attrib[_tag]
            _id = self.types[ _tag]
            _tags.update({ _id : _value })
  
        # Image > href
        _tag = '{%s}%s' % (self.ns['xlink'],'href')
        if _dict.__contains__(_tag):
            _value = _child.attrib[_tag]
            _id = self.types[ _tag]
            _tags.update({ _id : _value })
 
        #else:
        #    _tags.update({'level' : None} )
        
        return _tags

    def get_spaces( self, _child):
        """Deals with <text:s> which deals with extra spaces.
        """
        _tag = '{%s}%s' % (self.ns['text'],'c')
        if _child.attrib.__contains__(_tag):
            _value = _child.attrib[_tag]
            _tmp = ' ' * int(_value)
        else:
            _tmp = ' '
        return _tmp

    def get_tabs( self, _child):
        """Deals with <text:tab> which deals with extra spaces.
        """
        _tag = '{%s}%s' % (self.ns['text'],'tab')
        _tmp = '    '
        return _tmp

    
    def __fileExtractor__(self,_filename):
        '''Extract the file an assigns a proper filename to it'''
        self.files.extract( _filename,self.tmpdir )

        _extractedfile = os.path.join(self.tmpdir,_filename)

        # Is an OLE object.
        _toRename = False
        _ole = None
        if ole.isOleFile(_extractedfile):
            _ole = Ole( filename = _extractedfile )
            _new = _ole.extractFile()
            _ole.__close__()
            _extractedfile = _new
            _toRename = True
        
        if _toRename:
            _extractedfile = os.path.realpath(_extractedfile)
            _tmp = os.path.split(_new)
            _tmpname = _tmp[1]
            _tmpname = _tmpname.replace(' ','_')
            _new = os.path.join(_tmp[0],_tmpname )
        
            if self.isOverWriter:
                os.remove( _new )
            try:
                os.rename(_extractedfile,_new)
            except WindowsError:
                print 'ERROR: file already exists: %s' % _new
            _extractedfile = _new
        return _extractedfile
        # Los ficheros se extraen a sus rutas originales
        # Crea directorios si fuera necesario.
        # - Movemos los ficheros al mismo directorio que el .org.
        # - Usamos como prefijo el nombre del fichero.
        #_tmp = os.path.split( _extractedfile )
        #_newname = self.prefix + _tmp[1]
        #_fullnewname = os.path.join(_outdir,_newname)
        
        #if os.path.isfile( _fullnewname ) and self.isOverWriter:
        #    os.remove( _fullnewname )  # The preexisting file is removed.
        #elif os.path.isfile( _fullnewname ) and not self.isOverWriter:
        #    print 'WARNING - The file already exists.'
        #    print '  - FILENAME: %s' % _fullnewname
        #    print '  - Keeping both files'
        #    print '  - Referencing to the old one'
        #else:
        #    pass
        
        # We try to move the file changing its name. 
        #try:
        #    os.rename(_extractedfile,_fullnewname)
        #    return _newname
        #except:   
        #    return _filename

#=================================
class Ole:
    def __init__(self,filename=None):
        self.filename = filename
        self._data = ole.OleFileIO(self.filename)
        self.files = self.__getFiles__()
        self._dict = { 
             '.*AcroExch\.Document\.[0-9]+.*' : { 'ext':'.pdf','name':'CONTENTS'},
             '.*Word\.Document\.[0-9]+.*' : { 'ext':'.doc','name': None},
             '.*Excel\.Sheet\.[0-9]+.*'   : { 'ext':'.xls','name': None},
             '.*PowerPoint\.Show\.[0-9]+.*'   : { 'ext':'.ppt','name': None},
                     }
    
    def __guessFormat__(self):
        _data = self.__getItem__(0)
        _KEY = None
        for _key in self._dict.keys():
            _kernel = re.compile(  _key )
            _tmp = _kernel.findall( _data )
            if len(_tmp) == 1:
                _KEY = _key
        return _KEY

    def __getFiles__(self):
        return self._data.listdir()
    
    def __getItem__(self,_idx):
        _tmp = self._data.openstream( self.files[_idx] )
        _data = _tmp.read()
        return _data
        
    def __extractFile__(self,_idx):
        # Leemos el fichero
        _data = self.__getItem__(_idx)
        _fp = open('filaneme.txt','wb')
        _fp.write( _data )
        _fp.close()
    
    def extract_file(self):
        _key = self.__guessFormat__()
        if _key == None:
            _tmp = self._data.openstream( self.files[0] )
            _kernel = re.compile('[a-zA-Z]+\.[a-zA-Z]+\.[0-9]+')
            _tmp = _tmp.read()
            print 'ERROR: It is not configured for: %s' % _kernel.findall( _tmp )[0]
            #print self.files
            #print _tmp
            return None
        else:
            if self._dict[_key]['name'] != None:
                _tmp = self._data.openstream( self._dict[_key]['name'] )
                _file = os.path.splitext(self.filename)
                _ext = self._dict[ _key ]['ext']
                _filename = _file[0] + _ext
                #_tmp = self.outfile +_ext
                #_tmp = os.path.join( self.outdir, _tmp)
                _fp = open( _filename, 'wb')
                _fp.write( _tmp.read() )
                _fp.close()
    #            print self.filename
                del(_tmp)
                self._data.fp.close()
                os.remove(self.filename)
                return _filename
            else:
                # Símplemente se renombra
                #print self.files
                #print self.__getItem__(7)
                _file = os.path.splitext(self.filename)
                _ext = self._dict[ _key ]['ext']
                _filename = _file[0] + _ext                
                self._data.fp.close()
                try:
                    os.rename(self.filename,_filename)
                except WindowsError:
                    print 'WARNING: the file already exists: %s' % _filename
                return _filename

    def __close__(self):
        self._data.fp.close()

class FileError(Exception):
    def __init__(self, 
                 message,
                 filename ):
        self.filename = filename
        self.message = message

    def __str__(self):
        return u'%s: %s' % ( self.message, self.filename )

if __name__ == '__main__':
    pass
