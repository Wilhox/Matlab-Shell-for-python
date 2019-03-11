"""
Matlab Shell v.1.0

- Use python to run your m-files and Matlab commands
- Directly run m-scipts at shared session: arg1 = filebaseName, arg2 = filePath


This is free software; you can redistribute it and/or modify
it. This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY;

(C) 2019 Ville Lauronen <ville.lauronen(at)gmail.com>
"""

INSTRUCTIONS = """
========================================================================
INSTRUCTIONS
1. Any Matlab command is valid, Matlab engine is running at background.
2. Move to desktop by using command "desktop"
3. Run a script by using command "run yourfile.m"
4. Show current dir by using command "dir"
5. Press Enter to list your variables
6. Type "help" or "doc" to get more information
========================================================================
"""

AUTOSETUP = '''
========================================================================
MATLAB API FOR PYTHON

REQUIREMENTS

1. Windows with Python 3.5 and Matlab R2015 or newer
2. 64 bit Matlab requires 64 bit python
3. 32 bit Matlab requires 32 bit python

https://se.mathworks.com/help/matlab/matlab_external/install-the-matlab-engine-for-python.html

Press enter to install, administrator rights are needed.
========================================================================
'''

SHELL = 'MatlabShell>>>'

def PackageInstall(error):
    """
    - Finds out which package is missing
    - Downloads it automatically after five seconds.
    - Restarts script
    - Example:
        try:
            import numpy as np
            import matplotlib.pyplot as plot
        except ImportError as error:
            PackageInstall(error)
    """
    import time, subprocess, os, sys
    module = str(error)[15:].replace('\'', '')
    print(__doc__)
    print('>>>',str(error))
    print('>>> Downloading missing modules, please wait...')
    print('>>> The scirpt may restart multiple times')
    if 'win32com'in module or 'win32api' in module: #win32com and win32api must be installed as pywin32
        module = 'pypiwin32'
    if subprocess.call("pip install " + module):
        input('Press enter to continue')
    time.sleep(1)
    os.startfile(__file__)
    sys.exit()

try:
    import sys, cmd, logging
except ImportError as e:
    PackageInstall(e)

#Installing matlab plugin
try:
    import matlab.engine
except ImportError:
    try:
        import os, ctypes, subprocess, win32api, win32con, win32event, win32process
        from win32com.shell.shell import ShellExecuteEx
        from win32com.shell import shellcon
    except ImportError as e:
        PackageInstall(e)

    try:
        #administrator rights are needed
        if ctypes.windll.shell32.IsUserAnAdmin():

            matpath = os.environ['PROGRAMW6432'] + '\\MATLAB'

            #find newest version
            year = 0
            version = 'b'
            for folder in os.listdir(matpath):
                folder = folder.replace('R', '')
                try:
                    if int(folder[:-1]) > year:
                        year = int(folder[:-1])
                        version = folder[-1]
                except:
                    pass

            matpath = matpath + '\\R' + str(year) + version
            print('>>> Matlab engine for python is being installed from ' + matpath+'...')
            matpath += '\\extern\\engines\\python'
            os.chdir(matpath)
            feedback = subprocess.getoutput("python  setup.py install")
            for rivi in feedback.splitlines():
                print('>>>',rivi)

            #restart
            input('Press enter to continue')
            os.startfile(__file__)
            sys.exit()

        else: #Ask for install and elevate rights
            print(AUTOSETUP)
            input()
            process = ShellExecuteEx(
                nShow = win32con.SW_SHOWNORMAL,
                fMask = shellcon.SEE_MASK_NOCLOSEPROCESS,
                lpVerb = 'runas',
                lpFile = sys.executable,
                lpParameters = __file__)
            raise SystemExit

    except FileNotFoundError:
        print('>>> MATLAB not found from ' + matpath)
        print('More information: https://se.mathworks.com/help/matlab/matlab_external/install-the-matlab-engine-for-python.html')
        input()
    except SystemExit:
        raise SystemExit
    except:
        logging.exception('Internal Error')
        input()



class MatlabShell(cmd.Cmd):
    prompt = '>>> '
    file = None

    def __init__(self, engine = None, completekey='tab', stdin=None, stdout=None):
        if stdin is not None:
            self.stdin = stdin
        else:
            self.stdin = sys.stdin
        if stdout is not None:
            self.stdout = stdout
        else:
            self.stdout = sys.stdout
            self.cmdqueue = []
            self.completekey = completekey

        if engine == None:
            try:
                print(SHELL,'Matlab engine is starting...')
                self.engine = matlab.engine.start_matlab()
                print(SHELL,'Matlab is running')
            except:
                logging.exception(SHELL,'MSTARTUP FAILED')
                input()
                sys.exit()
        else:
            self.engine = engine

        self.cmdloop()

    def do_run(self, line): #run m-file by using command "run filename.m"
        '''Run m-file by using command "run example.m"'''
        if line.endswith(')'):
            line = line.replace(')', '')
            line = line.replace('(', '')
        if line.endswith('.m'):
            line = line.replace('.m', '')
        try:
            getattr(self.engine, line)(nargout=0)
        except:
            pass

    def do_help(self, line):
        try:
            getattr(self.engine, 'help')(line, nargout=0)
        except:
            pass

    def do_doc(self, line):
        if line.endswith(')'):
            line = line.replace(')', '')
            line = line.replace('(', ' ')
        try:
            getattr(self.engine, 'doc')(line, nargout=0)
        except:
            pass

    def do_shell(self, line):
        """Pass command to a system shell when line begins with '!'"""
        pass

    def emptyline(self):
        '''Show all variables when pressing enter and line is empty'''
        try:
            getattr(self.engine, 'who')(nargout=0)
        except:
            pass       

    def default(self, line):
        '''Unknown commands are send to Matlab'''
        try:
            getattr(self.engine, 'eval')(line, nargout=0)
        except:
            pass

if __name__ == "__main__":
    print('Matlab Shell v.1.0')
    print(INSTRUCTIONS)
    if len(sys.argv) < 2: #The script is started without arguments, -> Just start the shell
        MatlabShell()
    else: #The script is started with argumets, run m-script before starting shell.
        try:
            base = sys.argv[1]
            f = True
            for a in matlab.engine.find_matlab():
                if base == a:
                    print(SHELL,'Connecting to Matlab session named "' + base + '"...')
                    engine = matlab.engine.connect_matlab(base)
                    f = False
                    break
            if f:
                print(SHELL,'Creating shared Matlab session named "' + base + '"...')
                engine = matlab.engine.start_matlab()
                getattr(engine, 'matlab.engine.shareEngine')(base, nargout=0)
                engine.cd(sys.argv[2])

            try:
                getattr(engine, 'run')(base, nargout=0)
            except:
                pass
            MatlabShell(engine)
        except SystemExit:
            sys.exit()
        except:
            logging.exception('Internal Error')
            input()
