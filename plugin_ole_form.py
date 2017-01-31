#!/usr/bin/env python

__description__ = '/f plugin for oledump.py'
__author__ = 'Vincent Brillault'
__version__ = '0.0.1'
__date__ = '2017/01/29'

"""
History:
  2017/01/29: start
  2017/01/31: Better parsing, integrated with /o
"""

from oleform import ExtendedStream, consume_FormControl, consume_MorphDataControl
import cStringIO

VARIABLES = []

class cOF(cPluginParent):
    macroOnly = False
    name = 'UserForm plugin'

    def __init__(self, name, stream, options):
        self.streamname = name
        self.stream = stream
        self.options = options
        self.ran = False

    def Analyze(self):
        global VARIABLES
        result = []
        if len(self.streamname) > 1 and self.streamname[-1] in ['f', 'o']:
            self.ran = True
            stream = ExtendedStream(cStringIO.StringIO(self.stream), self.streamname)
            if self.streamname[-1] == 'f':
                VARIABLES = list(consume_FormControl(stream))
                if self.options == '-d':
                    for var in VARIABLES:
                        result.append('name: {name}, tag: {tag}, id: {id}, type: {ClsidCacheIndex}'.format(**var))
            else:
                for var in VARIABLES:
                    if var['ClsidCacheIndex'] != 23:
                        result.append('UNSUPPORTED DATA TYPE for {name}: {ClsidCacheIndex}'.format(**var))
                        return
                    var['value'] = consume_MorphDataControl(stream)
                    result.append('value for {name}: {value}'.format(**var))
                VARIABLES = []
        return result

AddPlugin(cOF)
