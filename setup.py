#!/usr/bin/env python
"""
Installs SimulationVBA using pip, setuptools or distutils

To install this package, run:
    pip install -e .

Or:
    python setup.py install

Installation using pip is recommended, to create scripts to run simulation_vba
and vbashell from any directory.
"""






try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup



entry_points = {
    'console_scripts': [
        'simulation_vba=simulation_vba.vba_emu:main',
        'vbashell=simulation_vba.vbashell:main',
    ],
}


setup(
    name="simulation_vba",
    version="1.0.3",
    description=(
        "SimulationVBA (fork of ViperMonkey) is a VBA Emulation engine written "
        "in Python, designed to analyze and deobfuscate malicious VBA Macros "
        "contained in Microsoft Office files (Word, Excel, PowerPoint, "
        "Publisher, etc)."),
    install_requires=[
        'oletools >= 0.56.1',
        "olefile",
        "prettytable",
        "colorlog",
        "colorama",
        "pyparsing==2.2.0",
        "unidecode==1.2.0",
        "xlrd",
        'regex; platform_python_implementation!="PyPy" or platform_system!="Windows"',
    ],
    packages=["simulation_vba", "simulation_vba.core"],
    setup_requires=["pytest-runner"],
    tests_require=["pytest"],
    entry_points=entry_points,
    author="Philippe Lagadec",
    url="https://github.com/decalage2/ViperMonkey",
    license="BSD",
    download_url="https://github.com/decalage2/ViperMonkey",
)
