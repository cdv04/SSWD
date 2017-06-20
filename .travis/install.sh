# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-06-20T13:17:52+02:00
# @Email:  zackary.b@live.fr
# @Project: PyMENT-SSWD
# @Filename: install.sh
# @Last modified by:   gysco
# @Last modified time: 2017-06-20T15:01:36+02:00

#!/bin/bash

if [[ $TRAVIS_OS_NAME == 'osx' ]]; then
  brew install wxmac wxpython
else
  sudo apt-get install ibwebkitgtk-dev libjpeg-dev libtiff-dev libgtk2.0-dev libsdl1.2-dev libgstreamer-plugins-base0.10-dev freeglut3 freeglut3-dev libnotify-dev wx-common
fi
pip install -r requirements.txt
if [[ $TRAVIS_OS_NAME == 'osx' ]]; then
  pyinstaller pyment-sswd_mac.spec -n pyment-sswd_mac --distpath=./dist/mac
  hdiutil create dist/pyment-sswd_mac.dmg -srcfolder dist/mac/ -ov
  zip -r dist/pyment-sswd_mac.zip dist/mac/pyment-sswd_mac.app
else
  pyinstaller pyment/__main__.py -w -n pyment-sswd_unix --distpath=./dist/unix
  tar -czvf dist/pyment-sswd_unix.tar.gz dist/unix/
fi
