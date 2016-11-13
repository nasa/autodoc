var=$1
#pandoc About.md -o about.html
#pandoc instructions.md -o instructions.html
javac -d . -cp  ':lib/*' src/*.java

if [ "$var" == "jar" ]; then
  jar -cvfm autodoc-v1.6.jar MANIFEST.MF -C bin bin/* gov/ html/instructions.html img/jpl_logo.png html/about.html
  rm -rf gov
fi
#rm instructions.html
#rm about.html
