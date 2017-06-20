# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-06-20T13:17:52+02:00
# @Email:  zackary.b@live.fr
# @Project: PyMENT-SSWD
# @Filename: install.sh
# @Last modified by:   gysco
# @Last modified time: 2017-06-20T19:07:39+02:00

#!/bin/bash
if [[ $TRAVIS_OS_NAME == 'osx' ]]; then
  pyinstaller pyment-sswd_mac.spec -n pyment-sswd_mac --distpath=./dist/mac
  hdiutil create dist/pyment-sswd_mac.dmg -srcfolder dist/mac/ -ov
  zip -r dist/pyment-sswd_mac.zip dist/mac/pyment-sswd_mac.app
else
  pyinstaller pyment/__main__.py -w -n pyment-sswd_unix --distpath=./dist/unix
  zip -r dist/pyment-sswd_unix.zip dist/unix/pyment-sswd_unix
fi
