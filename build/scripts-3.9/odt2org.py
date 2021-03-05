#!/usr/bin/env python
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
import os
from orgfile import ORGfile
from odtfile import ODTfile

if __name__ == '__main__':
    if not ERROR:
        import optparse
        _usage = u'''ODT2ORG: This software converts <filename.odt> (Open Office) into <filename.org> (emacs org-mode).

   General usage: <python_with_path> %prog [<options>] <input_file> [<output_file>]

Some examples:
- Convert "filename.odt" into "outputfile.org":
   python.exe odt2org.py filename.odt outputfile.org

- Covert  "filename.odt" into "filename.org"
   python odt2org.py filename.odt

- Using paths:
   C:\\python26\\python.exe C:\\odt2org\\odt2org.py C:\\MyFiles\\filename.odt

- Forcing overwriting:
   python odt2org.py -f filename.odt

'''
        _parser = optparse.OptionParser(usage = _usage, 
                                        version = "%prog 1.0" )

        _parser.add_option( "-f", 
                           "--force", 
                           action = "store_true",
                           dest = "overwrite",
                           help = "overwrite the output file if it exists")
        _parser.add_option( "-o", 
                           "--original", 
                           action = "store_true",
                           dest = "original",
                           help = "it shows spaces/tabs/returns as in the original file")

        (_options, _args) = _parser.parse_args()

        _isOK = True
        if len(_args) > 0:
            _inputfile = os.path.realpath( _args[0] )
        else:
            _isOK = False
        
        if not _isOK:
            print '''Get the help by writing: 

   python odt2org.py -h

'''
        
        if not os.path.isfile( _inputfile):
            print u'ERROR: the input file does not exist: %s' % _inputfile
            _isOK = False

        # If the outputfile is not provided, it is used one based in the original.
        if len( _args ) > 1:
            _outputfile = _args[1]
            _outputfile = os.path.realpath( _outputfile )
        else:
            _tmp = os.path.splitext( _inputfile )[0]
            _outputfile = os.path.realpath( _tmp + '.org' )
            
        # Check if output directory is valid.
        if not os.path.isdir( os.path.split(_outputfile)[0] ):
            print 'ERROR: output dir does not exists: %s' % os.path.split(_outputfile)[0]
            _isOK = False

        # Overwritting
        if not _options.overwrite and os.path.isfile( _outputfile ):
            print u'ERROR: output file already exist: %s' % _outputfile
            print u'       Use -f or --force to override'
            _isOK = False

        # We have proper inputfile, outputfile and overwrite option.
        if _isOK:
            _odtfile = ODTfile(filename = _inputfile)#,
                               #output = _outputfile,
                               #overwrite = _options.overwrite,
                               #original = _options.original)
            _list, _extra = _odtfile.gen_list()
            #overwrite = False
            _orgfile = ORGfile( filename = os.path.realpath( _outputfile ),
                                original = _inputfile,
                                overwrite = _options.overwrite)
            _orgfile.read_list( _list,_extra )
            _orgfile.export( )

       
