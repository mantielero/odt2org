#!/usr/bin/python
#-*- coding:utf-8 -*-

from distutils.core import setup

setup(name='odt2org',
      version='1.1',
      description='Converts from OpenOffice .odt files to Emacs org-mode files',
      classifiers=[
          'Development Status :: 5 - Production/Stable',
          'Environment :: Console',
          'Intended Audience :: End Users/Desktop',
          'Intended Audience :: Developers',
          'Intended Audience :: System Administrators',
          'License :: OSI Approved :: Apache Software License',
          'Natural Language :: English',
          'Operating System :: Microsoft :: Windows',
          'Operating System :: POSIX',
          'Programming Language :: Python',
          'Topic :: Documentation',
          'Topic :: Office/Business :: Office Suites',
          'Topic :: Text Processing',
          ],

      author= u'José María García Pérez',
      author_email= u'josemaria.alkala@gmail.com',
      url='http://bitbucket.org/josemaria.alkala/odt2org/wiki/Home',
      py_modules=['odtfile','orgfile'],
      scripts = ['./odt2org.py','odt2org.bat']
      )
