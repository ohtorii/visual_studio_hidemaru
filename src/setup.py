from distutils.core import setup
import py2exe

py2exe_options = {
    "unbuffered":True,
    "compressed": 1,
    "optimize": 2,
    "bundle_files": 1,
    'dll_excludes': [ "mswsock.dll", "powrprof.dll" ]
}

setup(  options = {"py2exe": py2exe_options},
        console = [
            {"script" : "visual_studio_hidemaru.py",}
        ],
        zipfile = None
    )
