### Run app:
`pipenv run python combine_word.py`


### Install EXE Converter:
`pipenv install pyinstaller`


### Create .exe:
1.  `pipenv run pyinstaller --onefile --noconsole --name CombineWord combine_word.py`
2.  `pipenv run pyinstaller --onefile --noconsole --name CombineWord --icon "icon.ico" combine_word.py`