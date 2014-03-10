#!/usr/bin/env python
from distutils.core import setup
import py2exe

includes = ["encodings","encodings.*"]
py2exe_opts = dict(
            ascii = True,
            compressed = True,
            optimize = 0,
            bundle_files = 1,
            excludes = ['_ssl','pyreadline','difflib','doctest','locale','optparse','pickle','calendar']
)

setup(
version = '1.0.0',
description = "PMS data validate",
author = 'wgzhao <wgzhao@gmail.com>',
options = {'py2exe':py2exe_opts},
zipfiile = None,
windows = [{'script':'main.py'}]
)