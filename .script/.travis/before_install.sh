# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-06-20T13:17:52+02:00
# @Email:  zackary.b@live.fr
# @Project: PyMENT-SSWD
# @Filename: install.sh
# @Last modified by:   gysco
# @Last modified time: 2017-06-20T19:07:29+02:00

#!/bin/bash

if [[ $TRAVIS_OS_NAME == 'osx' ]]; then
  brew install wxmac wxpython python3
  pip3 install https://github.com/pyinstaller/pyinstaller/archive/develop.zip
fi
