from distutils.core import setup
from Cython.Build import cythonize
setup(
    ext_modules = cythonize(["DownloadsSimplified.pyx", 
                             "DownloadsSimplifiednostartup.pyx",
                             "revertsorting.pyx"])
)