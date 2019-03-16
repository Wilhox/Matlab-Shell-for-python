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
6. Type "help" or "doc" to see Matlab help or documentation
7. Type "!help()" to access Matlab Shell documentation
========================================================================
"""

SYSTEMCOMMANDS = """SHELL COMMANDS

1. System commands always starts with "!"
2. Examples:
    - Stop matlab engine:
        >>> !engine.quit()

    - Start matlab engine
        >>> !startEngine()

    - Connect to matlab engine
        >>> sharedEngine(path, name):

    - Start socket;
        >>> !startSocket(port)

Press enter to browse all available system commands
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

PACKAGEINSTALL = '''
========================================================================
PYTHON PACKAGE INSTALLATION

REQUIREMENTS

1. Internet Connection
2. PIP

The script may restart multiple times, please wait...
========================================================================
'''

SHELL = 'MatlabShell>>> '
SOCKET = 'Socket>>> '
SERVER = 'Server>>> '

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
    print(PACKAGEINSTALL)
    if 'win32com'in module or 'win32api' in module: #win32com and win32api must be installed as pywin32
        module = 'pypiwin32'
    if subprocess.call("pip install " + module):
        input('Press enter to continue')
    time.sleep(1)
    os.startfile(__file__)
    sys.exit()

try:
    import sys, os, cmd, logging, time, threading, socket, io, json
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
    '''Main class for command shell'''
    prompt = '>>> '

    socketActive = False
    socketConnected = False
    matout = []
    materr = []


    def __init__(self, path = None, file = None, port = None, engine = None, completekey='tab', stdin=None, stdout=None):
        print('Matlab Shell v.1.0')
        print(INSTRUCTIONS)
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

        if port != None:
            self.startSocket(int(port))

        if path != None:
            self.path = path
        else:
            self.path = os.getcwd()

        if engine == None:
            self.startEngine()
        else:
            self.sharedEngine(self.path, engine)

        if file != None:
            if port != None:
                socketActive = True
                socketConnected = True
            self.run(file, self.path)

        self.cmdloop()


    def print(self, line, user = SHELL):
        '''print text'''
        self.stdout.write('\r'+user+str(line)+'\n')
        self.stdout.flush()

    def showPrompt(self):
        '''Erase line and print prompt'''
        self.stdout.write('\r'+self.prompt)
        self.stdout.flush()

    def startSocket(self, port):
        '''Start socket, agr1 = port'''
        try:
            self.socketActive = True
            host = socket.gethostname()
            self.connection = socket.socket()
            self.connection.bind((host, int(port)))
            self.socket =  threading.Thread(target = self.updateSocket)
            self.socket.start()
            self.print('Socket is running at ' + str(host)+'/'+str(port), user = SOCKET)
        except:
            logging.exception('SocketError')

    def stopSocket(self):
        self.socketActive = False

    def updateSocket(self):
        '''Socket update function running at own thread called "socket"'''
        while self.socketActive:
            try:
                time.sleep(1)
                self.connection.listen(1)
                c, adress = self.connection.accept()
                self.socketConnected = True

                while self.socketActive:
                    data = c.recv(1024).decode('utf-8')
                    if not data:
                        break
                    else:
                        self.print('Working...', user = SOCKET)
                        self.onecmd(data)
                        data = json.dumps([self.matout, self.materr])
                        c.send(data.encode('utf-8'))
                        self.matout.clear()
                        self.materr.clear()

                self.showPrompt()
                self.socketConnected = False
                c.close()
            except ConnectionResetError:
                pass

        self.connection.close()
        self.print('Stopped', user = SOCKET)

    def sharedEngine(self, path, name):
        '''
        - Connect to shared matlab session
        - Create new shared session if engine is not running
        - Args: arg1 = EnginePath, arg2 = EngineName, 
        '''
        try:
            f = True
            for a in matlab.engine.find_matlab():
                if name == a:
                    self.print('Connecting to Matlab engine named "' + name + '"...')
                    self.engine = matlab.engine.connect_matlab(name)
                    f = False
                    break
            if f:
                self.print('Creating shared Matlab engine named "' + name + '"...')
                self.engine = matlab.engine.start_matlab()
                getattr(self.engine, 'matlab.engine.shareEngine')(name, nargout=0)
        except:
            logging.exception('EngineError')

    def startEngine(self):
        '''
        Create Matlab engine
        args: None
        return: None
        '''
        try:
            self.print('Matlab engine is starting...')
            self.engine = matlab.engine.start_matlab()
        except:
            logging.exception(SHELL,'EngineError')

    def help(self):
        '''
        Show matlabShell documentation
        args: None
        return: None
        '''
        print(INSTRUCTIONS)
        print(SYSTEMCOMMANDS)
        try:
            for a in dir(self.__class__):
                if not a.startswith('__'):
                    input('\n'+50*'='+'\n')
                    print(str(a))
                    print('\n\n'+str(getattr(self, a).__doc__))
            input('\n'+50*'='+'\n')
        except:
            logging.exception('')

    def run(self, fileName, filePath = None):
        '''
        Run m-file by using matlab.run
        args: filename, path = None
        return: None
        '''
        if filePath != None and filePath != self.path:
            self.engine.cd(filePath)
            self.path = filePath
            os.chdir(filePath)
            self.print(filePath)
        try:
            self.evalMatlab('clear', fileName)#Refresh file by cleaning cache
            self.evalMatlab('run', fileName)
        except FileNotFoundError as e:
            self.ErrorHandler(e)
        except:
            logging.exception('InternalError')

    def eval(self, fileName):
        '''
        Run m-file by using matlab.eval
        args: filename
        return: None
        '''
        try:
            file = io.open(fileName+'.m')
            data = file.read()
            file.close()
            self.evalMatlab('eval', data)
        except FileNotFoundError as e:
            self.ErrorHandler(e)
        except:
            logging.exception('InternalError')

    def do_run(self, line):
        '''
        - Call matlab.run
        - Parse args before passing them to matlab
        args: line
        return: None
        '''
        if line.endswith(')'):
            line = line.replace(')', '')
            line = line.replace('(', '')
        if line.endswith('.m'):
            line = line.replace('.m', '')
        self.evalMatlab('clear', fileName)
        self.evalMatlab('run', line)

    def do_help(self, line):
        '''
        - Call matlab.help
        - Parse args before passing them to matlab
        args: line
        return: None
        '''
        if line.endswith(')'):
            line = line.replace(')', '')
            line = line.replace('(', ' ')
        self.evalMatlab('help', line)

    def do_doc(self, line):
        '''
        - Call matlab.doc
        - Parse args before passing them to matlab
        args: line
        return: None
        '''
        if line.endswith(')'):
            line = line.replace(')', '')
            line = line.replace('(', ' ')
        self.evalMatlab('doc', line)

    def do_shell(self, line):
        """
        Pass command to a system shell when line begins with '!'
        args: line
        return: None
        """
        try:
            if line.endswith(')'):
                line = line.split('(')
                if line[1].replace(')', '') != '':
                    self.parseCommand(line[0], line[1].replace(')', ''))
                else:
                    self.parseCommand(line[0], [])
            else:
                self.parseCommand(line)
        except AttributeError as e:
            self.ErrorHandler(e)
        except:
            logging.exception('InputError')

    def parseCommand(self, base, args = None):
        '''helps do_shell function'''
        command = self
        commands = base.split('.')
        nro = 0
        for a in commands:
            command = getattr(command, commands[nro])
            nro += 1

        if args != None:
            if args == []:
                command()
            else:
                command(*args.split(','))
        else:
            print(command)


    def emptyline(self):
        '''
        Show all variables when pressing enter and line is empty
        args: None
        return: None
        '''
        self.evalMatlab('who', '')  

    def default(self, line):
        '''
        - Unknown commands are send to Matlab.
        - Line is evaluated by using matlab.eval-function.

        args: None
        return: None
        '''
        self.evalMatlab('eval', line)

    def evalMatlab(self, command, args):
        '''
        Run command in matlab engine and get results.

        Example: Run Matlab command run file.m by using command
        evalMatlab('run', file)

        args: command, args
        return: None
        '''
        try:
            stdout = io.StringIO()
            stderr = io.StringIO()
            exe = getattr(self.engine, command)(args, stdout = stdout, stderr = stderr, nargout = 0, async = True)
            tellout = stdout.seek(0)
            tellerr = stderr.seek(0)
            done = False
            while not done:
                exe.result()
                tellout = self.captureOutput(tellout, stdout, self.matout)
                tellerr = self.captureOutput(tellerr, stderr, self.materr)
                self.stdout.flush()
                done = exe.done()
            stdout.close()
            stderr.close()

        except SyntaxError as e:
            self.ErrorHandler(e,'')
        except matlab.engine.MatlabExecutionError as e:
            self.ErrorHandler(e,'Matlab: ')
        except matlab.engine.RejectedExecutionError as e:
            self.ErrorHandler(e,'')
        except:
            logging.exception('')

    def ErrorHandler(self, error, user = None):
        '''
        - Print error
        - Save error if socket active
        '''
        if user == None:
            user = self.SHELL
        self.print(error, user)
        if self.socketConnected:
            self.materr.append(str(error))

    def captureOutput(self, tell, stream, capturelist):
        '''helper for function called "evalMatlab" '''
        stream.seek(tell)
        data = stream.readlines()
        if data != []:
            self.stdout.write('\n')
            for a in data:
                self.stdout.write(a)
            if self.socketConnected:
                capturelist += data
            self.stdout.write('\n')
            tell = stream.tell()+1
        return(tell)



if __name__ == "__main__":
    try:
        MatlabShell(*sys.argv[1:])
    except SystemExit:
        sys.exit()
    except:
        logging.exception('Internal Error')
        input()
