rm -rf Class\ Scanner.spec build/ dist/ ~/Library/Application\ Support/Class\ Scanner/config.yaml
pyinstaller \
  --windowed \
  --name "Class Scanner" \
  --icon app_icon.icns \
  --osx-bundle-identifier "ch.yourname.class-scanner" \
  --add-data "default_config.yaml:." \
  scan_app.py
