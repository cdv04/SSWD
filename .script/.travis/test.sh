# @Author: Zackary BEAUGELIN <gysco>
# @Date:   2017-06-20T13:17:52+02:00
# @Email:  zackary.b@live.fr
# @Project: PyMENT-SSWD
# @Filename: install.sh
# @Last modified by:   gysco
# @Last modified time: 2017-06-20T19:07:39+02:00

#!/bin/bash
if [[ $TRAVIS_OS_NAME == 'osx' ]]; then
  python3 -c "import wx; print(wx.__version__)"
  pyinstaller pyment-sswd_mac.spec -n pyment-sswd_mac --distpath ./dist/mac
  hdiutil create dist/pyment-sswd_mac.dmg -srcfolder dist/mac/ -ov
  zip -r dist/pyment-sswd_mac.zip dist/mac/pyment-sswd_mac.app
else
  LD_LIBRARY_PATH=/home/travis/virtualenv/$(python3 --version | awk '{ gsub (" ", "", $0); print tolower($0)}')/lib/python3.5/site-packages/wx/
  python3 -c "import wx; print(wx.__version__)"
  pyinstaller pyment-sswd_unix.spec -n pyment-sswd_unix --distpath ./dist/unix
  zip -r dist/pyment-sswd_unix.zip dist/unix/pyment-sswd_unix
fi
