# langcheck
A script to check if a given document contains Portuguese text.

Wiki needs to be updated.

Compilation notes (these steps necessary to ensure that package is small; otherwise, 200+ MB size)
1. Created virtual environment (virtualenv langcheck). Deactivate base anaconda environment
2. In virtual env, installed all libraries (xlrd, python-pptx, pypiwin32)
3. Steps for installing Polyglot on a windows system:
    a. Download whl files for PyICU and pycld2 from https://www.lfd.uci.edu/~gohlke/pythonlibs/. 
    Select the package matching your Python installation and windows.
    b. Do pip install for the whl files of PyICU and pycld2
    c. Download the polyglot master zip from https://github.com/aboSamoor/polyglot/archive/master.zip. Unzip the file.
    d. Navigate into the unziped folder and do python setup.py install
4. In virtual env, installed pyinstaller
5. Using pyinstaller -w -F (meaning not windowed and onefile), compiled script
