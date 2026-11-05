# Rebuild index.html from index.src.html (inlines the logo + icon).
s = open('index.src.html').read()
s = s.replace('/*LOGO*/', open('.logo.css').read().strip()).replace('<!--ICON-->', open('.icon.html').read().strip())
open('index.html', 'w').write(s)
