from __future__ import with_statement, print_function
# Main module for launching the standalone velocipy GUI
import inspect
import pickle
import subprocess
import os
import copy
import argparse
import sys
import traceback
import functools
import collections
import enum
import distutils.sysconfig

from win32com.client import Dispatch
import wx
import wx.aui
import wx.lib.agw.customtreectrl as CT
from wx.lib.mixins import treemixin
try:
    import wx.lib.iewin as iewin
    print("ActiveX")
except:
    import wx.html2 as webview
    print("HTML2")

from HSTB.shared import Constants
from HSTB.gui import BaseAuiFrame
from HSTB.resources import path_to_html, PathToDocs, PathToResource, path_to_HSTB, path_to_NOAA, path_to_NOAA_site_packages, create_env_cmd_list
from win32api import GetShortPathName as get_short_path_name
import win32api
import win32con

noaa_sitepkg_dir = path_to_NOAA("site-packages")
PathToSitePkgs = distutils.sysconfig.get_python_lib()

PathTo_hyo2 = path_to_NOAA_site_packages("Python3", "hyo2")

_dHSTP = Constants.UseDebug()  # Control debug stuff (=0 to hide debug menu et al from users in the field)
_dHSTP = True
_PydroVersion = Constants.PydroTitleVersion()
if not _dHSTP:
    # disable warnings; e.g.,
    # C:\Python23\lib\site-packages\wxPython\image.py:208: DeprecationWarning: integer argument expected, got float
    #  val = imagec.wx.Image_Scale(self, *_args, **_kwargs)
    def _theevilunwarner(*args, **kwargs):
        pass
    import warnings
    warnings.warn = _theevilunwarner
    warnings.warn_explicit = _theevilunwarner


# break the gui and data/operations into separate pieces as they should be.
# Now the Data portion can supply menu items for any gui frame to show as desired but it is not tied to a particular gui.
# Hence we can load the VelocipyData into a Pydro frame or a standalone lightweight app "Velocipy" with equivalent results.

RunTypeEnum = enum.IntEnum("RunType", (('PYTHON', 1), ('RAW', 2), ))
OptsEnum = enum.IntEnum('OptionNames', (("ARGS", 0), ("CMD", 1), ("ENV", 2), ("DIR", 3), ("CONSOLE", 4), ("DEBUG", 5),))
RTE = RunTypeEnum

ProgramList = {}

class ProgOpts(object):
    """ Parameters used to launch a program

        These are program start arguments to be passed to the CreateArgs function before launching or creating a startup icon
        - first is  list of script/program to run (normally script name passed to python.exe) and additional command line parameters
        - second value is the executable to use.
        - third value is the python environment (env) to use -- can be empty string
        - fourth value is the start directory, relative to {PydroInstallRoot}/NOAA/site-packages
        - fifth value is boolean of if a new console should be spawned.  Use for launching iPython shell or maybe some other programs.
        - sixth value is boolean of if the console should remain after the program exits. Use for debugging a program that has the console disappear.
            could use new shell as default -- would spawn a bunch of console windows potentially -- output would be separated but too many consoles?

        OptsEnum = enum.IntEnum('OptionNames', (("ARGS", 0), ("CMD", 1), ("ENV", 2), ("DIR", 3), ("CONSOLE", 4), ("DEBUG", 5),))

        NOTE when run from the application any arguments with spaces will get quotes around them
        but when run from an icon they will just be joined together with spaces.

        ex:  set pythonpath=test => "set pythonpath=test" in subprocess.popen fails but would have worked from the icon
        using: ["set", "pythonpath=test"] works in both icon and subprocess

    """

    def __init__(self, args=None, cmd="", env="", dir="", new_console=False, persist_console=False):
        self.dir = dir
        self.persist_console = persist_console
        self.new_console = new_console
        self.env = env
        self.cmd = cmd
        if args is None:
            args = []
        self.args = copy.copy(args)

    def copy(self):
        return ProgOpts(self.args, self.cmd, self.env, self.dir, self.new_console, self.persist_console)

class PythonOpts(ProgOpts):
    def __init__(self, args=None, env="", dir="", new_console=False, persist_console=False):
        super(PythonOpts, self).__init__(args=args, cmd=RTE.PYTHON, env=env, dir=dir, new_console=new_console, persist_console=persist_console)

class Program:
    def __init__(self, name, run_opts=[], docs=None, descr="", desktop_icon=None, tree_icon=None):
        """
        Parameters
        ----------
        name
            Name to be displayed in the navigation tree on the left side
        run_opts
            A ProgOpts instance or a list of parameters to be passed to ProgOpts(*run_opts)

            These are program start arguments to be passed to the CreateArgs function before launching or creating a startup icon
            - first is  list of script/program to run (normally script name passed to python.exe) and additional command line parameters
            - second value is the executable to use.
            - third value is the python environment (env) to use -- can be empty string
            - fourth value is the start directory, relative to {PydroInstallRoot}/NOAA/site-packages
            - fifth value is boolean of if a new console should be spawned.  Use for launching iPython shell or maybe some other programs.
            - sixth value is boolean of if the console should remain after the program exits. Use for debugging a program that has the console disappear.
                could use new shell as default -- would spawn a bunch of console windows potentially -- output would be separated but too many consoles?

            OptsEnum = enum.IntEnum('OptionNames', (("ARGS", 0), ("CMD", 1), ("ENV", 2), ("DIR", 3), ("CONSOLE", 4), ("DEBUG", 5),))

            NOTE when run from the application any arguments with spaces will get quotes around them
            but when run from an icon they will just be joined together with spaces.

            ex:  set pythonpath=test => "set pythonpath=test" in subprocess.popen fails but would have worked from the icon
            using: ["set", "pythonpath=test"] works in both icon and subprocess
        docs
            file path to the html or other docs to be displayed in the html window in Pydro Explorer
        descr
            One line description that will be used in tooltips and in the auto-generated html
        desktop_icon
            Icon to show in the taskbar or on an icon made for the start menu or desktop
        tree_icon
            min-icon to be displayed in the tree on the left of Pydro Explorer.
            Can be None or "recent" or a path to an image.
        """
        self._name = None
        if desktop_icon is None:
            icon = PathToResource("Pydro.ico")
        if tree_icon == "recent":
            tree_icon = PathToResource("recent.png")
        if docs is None:
            docs = path_to_html("Pydro", "General.html")
        if not descr:
            descr = "{} didn't have a documentation entry".format(name)
        self.name = name
        self.descr = descr
        if isinstance(run_opts, ProgOpts):
            self.opts = run_opts
        else:
            self.opts = ProgOpts(*run_opts)
        self.docs = docs
        self.tree_icon = tree_icon
        self.desktop_icon = desktop_icon
        ProgramList[self.name] = self
    @property
    def name(self):
        return self._name
    @name.setter
    def name(self, val):
        if val in ProgramList:  # make sure the new name doesn't already exist
            raise Exception("%s Duplicate program name"%val)
        try:  # remove the old name - if there was one
            ProgramList.pop(self._name)
        except KeyError:
            pass
        self._name = val
        ProgramList[self._name] = self  # add to the master program list

download_aviso = Program("Download Aviso FES Tide Data",
                         ProgOpts(cmd=path_to_NOAA_site_packages("aviso.bat"), persist_console=True),
                         path_to_html("Pydro", "General.html"),
                         "Data to support global tide estimates")

download_gebco_gdb = Program("Download GEBCO Data for ArcGIS / Beets",
                             ProgOpts(cmd=path_to_NOAA_site_packages("gebco_gdb.bat"), persist_console=True),
                             path_to_html("Beets", "Beets_Documentation.html"),
                         "Data to support the 'Beets' survey effort estimates")

beets = Program("Beets",
                # ProgOpts(),  # nothing to run
                docs=path_to_html("Beets", "Beets_Documentation.html"),
                descr="Data to support the 'Beets' survey effort estimates")

download_gebco_rasters = Program("Download GEBCO Data for external use",
                                 ProgOpts(cmd=path_to_NOAA_site_packages("gebco_rasters.bat"), persist_console=True),
                                 path_to_html("Beets", "Beets_Documentation.html"),
                         "Data to support the 'Beets' survey effort estimates")

lnm_calc = Program("List XmlDR Stats",
                   PythonOpts(["list_xmldr_stats.py", ], "Pydro367", "Python3\\HSTB\\scripts", True),
                   path_to_html("Pydro", "General.html"),
                   "Extract Linear Nautical Miles stats from XmlDR files"
                   )

general_docs = Program("General",
                       docs=path_to_html("Pydro", "General.html"),
                       descr="General documentation")

arcmap_docs = Program("ESRI-Arc",
                      descr="Connecting ESRI/Arc documentation",
                      docs=path_to_html("Arc", "connecting.html"))

HSRR = Program("HSRR helper",
                   ProgOpts([path_to_NOAA_site_packages("enablePyQt.bat"), "&&", "python", "HSRR_GUI.py", ], "", "Pydro367", "Python3\\HSTB\\HSRR_GUI", True),
                   path_to_NOAA_site_packages(r"Python3\HSTB\HSRR_GUI\HSRRHelper.html"),
                   "Hydrographic Systems Readiness Review (HSRR) helper"
                   )

files_checker = Program("Files Checker",
                   PythonOpts(["Files_checker.py", ], "Pydro367", "Python3\\HSTB\\Files_Checker", True),
                   path_to_NOAA_site_packages(r"Python3\HSTB\Files_Checker\Read_me.htm"),
                   "App to do something helpful :)"
                   )


ProgramEnum = enum.Enum('ProgramNames',
                        """TOGGLE HYPACKLINES BENCHMARK ENCPRODSPEC PHBCOPY S7K
                        IPYTHON IPYTHONWX IPYTHONQT IPYTHONNOTEBOOK PYTHONWIN SPYDER
                        IPYTHON27 IPYTHONWX27 IPYTHONQT27 IPYTHONNOTEBOOK27 PYTHONWIN27 SPYDER27
                        FETCHTIDES IDLE SATMON ROOMBA DIR_SIZES
                        VELOCIPY AUVDEPTH CASTTIME AUTOQC CSARQA BDB_ASCII
                        OPENBST SIS4 SIS5 SOUNDSPEED HDF_COMPASS STORMFIX
                        QCTOOLS CATOOLS QAX ENCX BAGEXPLORER FIGLEAF BRESS VDATUM_SEP PYDROGIS XMLDR POSTACQ
                        MAKECATALOG CHARLENE S57COMPARE ACQFILETRANSFER SHAM SCRIBBLE SIMPLE_TCARI SIMPLE_TIDES_REQ GRIDCOMP NCEICHECK
                        LICENSES27 LICENSES LTD CONSOLE27 CREATE38ENV CONSOLE38 SPYDER38 CONSOLE367 DEMONSTRATOR27 DEMONSTRATOR36
                        PERMISSIONS SCRIPT_FLIERS SCRIPT_UNCERTAINTY SURVEY_OUTLINES VR_BAG
                        IMAGE_RENAME WEEKLYREP SEPERATE_2040_710_FREQ WXDEMO27 WXDEMO36
                        NOAA_S57 PYTHON_BASICS REVERT_PB_NOTEBOOKS OCEAN_DATA_SCIENCE REVERT_ODS_NOTEBOOKS
                        TJ_ACQ_LOG NBS_EMAIL PROD_EMAIL SHPO_EMAIL PICKY
                        SUSSIE
                        """)

ProgramNames = {ProgramEnum.SATMON: 'Satmon',
                ProgramEnum.ROOMBA: 'Roomba',
                ProgramEnum.VELOCIPY: 'Velocipy',
                ProgramEnum.LTD: 'Transmission Letter',
                ProgramEnum.AUVDEPTH: 'AUV Depth',
                ProgramEnum.CASTTIME: 'CastTime',
                ProgramEnum.AUTOQC: 'POSPacAutoQC',
                ProgramEnum.CSARQA: 'Finalized CSAR QA',
                ProgramEnum.BDB_ASCII: 'BDB Surface ASCII Export Stats',
                ProgramEnum.BRESS: 'BRESS',
                ProgramEnum.OPENBST: 'OpenBST',
                ProgramEnum.SIS4: 'SIS4 Emulator',
                ProgramEnum.SIS5: 'KCtrl Emulator',
                ProgramEnum.SOUNDSPEED: 'Sound Speed Manager',
                ProgramEnum.HDF_COMPASS: 'HDF Compass',
                ProgramEnum.STORMFIX: 'StormFix',
                ProgramEnum.QCTOOLS: 'QC Tools',
                ProgramEnum.CATOOLS: 'CA Tools',
                ProgramEnum.QAX: 'QAX',
                ProgramEnum.ENCX: 'ENC X',
                ProgramEnum.BAGEXPLORER: 'BAG Explorer',
                ProgramEnum.FIGLEAF: 'FigLeaf',
                ProgramEnum.VDATUM_SEP: 'VDatum SEP from Shapefile',
                ProgramEnum.PYDROGIS: 'PydroGIS',
                ProgramEnum.XMLDR: 'XmlDR',
                ProgramEnum.POSTACQ: 'PostAcquisitionTools',
                ProgramEnum.PYTHONWIN: "PythonWin (Python 3.6)",
                ProgramEnum.PYTHONWIN27: "PythonWin (Python 2.7)",
                ProgramEnum.SPYDER: "Spyder (Python 3.6)",
                ProgramEnum.SPYDER38: "Spyder (Python 3.8)",
                ProgramEnum.SPYDER27: "Spyder (Python 2.7)",
                ProgramEnum.IDLE: "IDLE",
                ProgramEnum.FETCHTIDES: "FetchTides",
                ProgramEnum.CREATE38ENV: "Create the Python3.8 Testing Environment",
                ProgramEnum.CONSOLE38: "Python ready console (Python 3.8.1 2020)",
                ProgramEnum.CONSOLE367: "Python ready console (Python 3.6.7)",
                ProgramEnum.IPYTHON: "IPython (Python 3.6)",
                ProgramEnum.IPYTHONWX: "wx IPython (Python 3.6)",
                ProgramEnum.IPYTHONQT: "qt IPython (Python 3.6)",
                ProgramEnum.IPYTHONNOTEBOOK: "Jupyter (IPython) Notebook (Python 3.6)",
                ProgramEnum.CONSOLE27: "Python ready console (Python 2.7)",
                ProgramEnum.IPYTHON27: "IPython (Python 2.7)",
                ProgramEnum.IPYTHONWX27: "wx IPython (Python 2.7)",
                ProgramEnum.IPYTHONQT27: "qt IPython (Python 2.7)",
                ProgramEnum.IPYTHONNOTEBOOK27: "Jupyter (IPython) Notebook (Python 2.7)",
                ProgramEnum.S7K: "7K to S7K",
                ProgramEnum.PHBCOPY: "PHB QuickTransfer",
                ProgramEnum.MAKECATALOG: "Make 000 Catalog",
                ProgramEnum.ENCPRODSPEC: "Change ENC Product Spec",
                ProgramEnum.BENCHMARK: "Caris Performance Benchmark",
                ProgramEnum.WEEKLYREP: "Weekly Reports",
                ProgramEnum.HYPACKLINES: "ArcMap Lines for Hypack",
                ProgramEnum.DEMONSTRATOR27: "Common Code Base Explorer (Python 2.7)",
                ProgramEnum.DEMONSTRATOR36: "Common Code Base Explorer (Python 3.6)",
                ProgramEnum.WXDEMO27: "wxPython Demo (Python 2.7)",
                ProgramEnum.WXDEMO36: "wxPython Demo (Python 3.6)",
                ProgramEnum.CHARLENE: 'Charlene',
                ProgramEnum.S57COMPARE: 'S57 Compare',
                ProgramEnum.ACQFILETRANSFER: 'Acquisition File Transfer',
                ProgramEnum.SHAM: 'Sham (Shoreline attribution)',
                ProgramEnum.SCRIBBLE: 'Scribble (Automated Reports)',
                ProgramEnum.TOGGLE: "Toggle Auto-Updates",
                ProgramEnum.SIMPLE_TCARI: "Apply TCARI",
                ProgramEnum.SIMPLE_TIDES_REQ: "Tides Request",
                ProgramEnum.GRIDCOMP: "Compare Grids",
                ProgramEnum.NCEICHECK: "NCEI Checkout",
                ProgramEnum.LICENSES27: "License Information (Python27)",
                ProgramEnum.LICENSES: "License Information (Python36)",
                ProgramEnum.PERMISSIONS: "Fix File Permissions",
                ProgramEnum.SURVEY_OUTLINES: "Extract Survey Outlines",
                ProgramEnum.SCRIPT_FLIERS: "Script to Find Fliers",
                ProgramEnum.SCRIPT_UNCERTAINTY: "Script for Empty Uncertainty",
                ProgramEnum.VR_BAG: "VR to SR Bag",
                ProgramEnum.IMAGE_RENAME: "Rename FFF Images per HTD",
                ProgramEnum.NBS_EMAIL: "NBS Mass Transmittal Email",
                ProgramEnum.PROD_EMAIL: "HSD Digital Production Transmittal Email",
                ProgramEnum.SHPO_EMAIL: "SHPO Email",
                ProgramEnum.DIR_SIZES: "Check directory sizes and report in CSV",
                ProgramEnum.SEPERATE_2040_710_FREQ: "Separate EM2040 and EM710 by frequency",
                ProgramEnum.NOAA_S57: "NOAA S57 Support Files",
                ProgramEnum.PYTHON_BASICS: "Open Programming Basics with Python",
                ProgramEnum.REVERT_PB_NOTEBOOKS: "Reset Programming Basics with Python",
                ProgramEnum.OCEAN_DATA_SCIENCE: "Open Introduction to Ocean Data Science",
                ProgramEnum.REVERT_ODS_NOTEBOOKS: "Reset Introduction to Ocean Data Science",
                ProgramEnum.TJ_ACQ_LOG: "Acquisition Log",
                ProgramEnum.PICKY: "Picky",
                ProgramEnum.SUSSIE: 'Sussie',
                }
PN = ProgramNames
PE = ProgramEnum

shell = Dispatch('WScript.Shell')
docs_path = shell.SpecialFolders("MyDocuments")
# These are program start arguments to be passed to the CreateArgs function before launching or creating a startup icon
# first is  list of script/program to run (normally script name passed to python.exe) and additional command line parameters
# second value is the executable to use.
# third value is the python environment (env) to use -- can be empty string
# fourth value is the start directory, relative to {PydroInstallRoot}/NOAA/site-packages
# fifth value is boolean of if a new console should be spawned.  Use for launching iPython shell or maybe some other programs.
# sixth value is boolean of if the console should remain after the program exits. Use for debugging a program that has the console disappear.
#   could use new shell as default -- would spawn a bunch of console windows potentially -- output would be separated but too many consoles?


# NOTE when run from the application any arguments with spaces will get quotes around them
# but when run from an icon they will just be joined together with spaces.
# ex:  set pythonpath=test => "set pythonpath=test" in subprocess.popen fails but would have worked from the icon
# using: ["set", "pythonpath=test"] works in both icon and subprocess
jupyter_docs = docs_path if "JUPYTER_PATH" not in os.environ else os.environ["JUPYTER_PATH"]
ProgramOpts = {
    PN[PE.XMLDR]: PythonOpts(["-m", r"HSTB.gui.xmlDR"], "Pydro27"),
    PN[PE.SHAM]: PythonOpts(["velodyne_csv_to_s57.py", ], "Pydro27", "Python2\\HSTB\\Charlene"),
    PN[PE.SCRIBBLE]: PythonOpts(["dr_dump.py", ], "Pydro27", "Python2\\HSTB\\Charlene"),
    PN[PE.CASTTIME]: [["--pylab=wx", "StartModule.py", "CastTimeGui"], "ipython.exe", "Pydro27", "Python2\\HSTP\\Pydro", ],
    PN[PE.CHARLENE]: [["charlene.py", ], RTE.PYTHON, "Pydro27", "Python2\\HSTB\\Charlene", ],
    PN[PE.S57COMPARE]: [["s57compare_gui.py", ], RTE.PYTHON, "Pydro27", "Python2\\HSTB\\s57compare", ],
    PN[PE.ACQFILETRANSFER]: [["Acq_transfer.py", ], RTE.PYTHON, "Pydro367", "Python3\\HSTB\\Acq_file_transfer", ],
    PN[PE.SATMON]: [["StartModule.py", r"satmon"], RTE.PYTHON, "Pydro27", "Python2\\HSTP\\Pydro", ],
    PN[PE.ROOMBA]: [["-m", "HSTB.gui.roomba"], RTE.PYTHON, "Pydro367", ],
    PN[PE.PYDROGIS]: [["StartModule.py", r"Pydro"], RTE.PYTHON, "Pydro27", "Python2\\HSTP\\Pydro", ],
    PN[PE.POSTACQ]: [["StartModule.py", r"PostAcquisitionTools"], RTE.PYTHON, "Pydro27", "Python2\\HSTP\\Pydro", ],
    PN[PE.TJ_ACQ_LOG]: [["-m", "HSTB.acq_log"], RTE.PYTHON, "Pydro367"],
    PN[PE.IDLE]: [["/c " + PathToSitePkgs + "\\..\\idlelib\\idle.bat", ], 'cmd.exe', "Pydro27", "Python2\\HSTP\\Pydro", ],
    PN[PE.S7K]: [["Pydro7K2s7K.py", ], RTE.PYTHON, "Pydro27", "Python2\\HSTP\\Pydro\\Macros", ],
    PN[PE.BENCHMARK]: [["CarisBenchmarking27_V2.py", ], RTE.PYTHON, "Pydro27", "Python2\\HSTP\\Contribs\\CarisBenchmark", ],
    PN[PE.WEEKLYREP]: [[], RTE.PYTHON, None],
    PN[PE.HYPACKLINES]: [[], None, None],
    PN[PE.TOGGLE]: [["CheckForUpdates.py", "-TOGGLE"], RTE.RAW, "Pydro27", "Python2\\HSTP\\Pydro", ],

    PN[PE.LTD]: [["-m", "HSTB.gui.datatransfer"], RTE.PYTHON, "Pydro367"],
    PN[PE.PYTHONWIN]: [["Pydro367"], path_to_NOAA_site_packages("run_pythonwin.bat"), "base", "", True],
    # PN[PE.PYTHONWIN]: [[], PathToSitePkgs.lower().replace("\\envs\\pydro27", "\\envs\\pydro367") + '\\pythonwin\\Pythonwin.exe', "Pydro367"],
    PN[PE.CONSOLE367]: [[], "", "Pydro367", "Python3", True, True],
    PN[PE.SPYDER38]: [[], "spyder", "Pydro38_Test", "", True],
    PN[PE.CREATE38ENV]: [[], path_to_NOAA_site_packages("MakePydro38_TestEnv.bat"), "", "", True, True],
    PN[PE.CONSOLE38]: [[], "", "Pydro38_Test", "Python38", True, True],
    PN[PE.IPYTHON]: [["--ipython-dir=%s" % docs_path], "ipython.exe", "Pydro367", "", True, True],
    PN[PE.IPYTHONWX]: [["--pylab=wx", "--ipython-dir=%s" % docs_path], "ipython.exe", "Pydro367", "", True, True],
    PN[PE.IPYTHONQT]: [["--pylab=qt", "--ipython-dir=%s" % docs_path], "ipython.exe", "Pydro367", "", True, True],
    PN[PE.IPYTHONNOTEBOOK]: [["notebook", "--notebook-dir=%s" % jupyter_docs], "jupyter", "Pydro367", "", True, True],
    #    PN[PE.SPYDER]: [[], "spyder", "Pydro367", "", True],
    # PN[PE.SPYDER]: [[path_to_HSTB("..\..\enablePyQt.bat"), "&&", path_to_HSTB(r"..\..\RunSpyder36_2019.bat")], "", "Pydro367", "", True],
    PN[PE.SPYDER]: [[], path_to_NOAA_site_packages("RunSpyder36_2019.bat"), "Pydro367", "", True],
    # Setting the python path to the Python27 modules lets the demo code run without making a second copy in the Python3 directory.
    # There can't be spaces in the pythonpath so strip any spaces off the pkg_dir and then split it to make params without spaces.
    # Conda doesn't allow spaces so the pkg_dir.split(" ") isn't really necessary
    # if there were spaces in the path it should work though due to strip() and split()
    PN[PE.DEMONSTRATOR36]: [['-m', 'HSTB.gui.demo'], RTE.PYTHON, "Pydro367"],
    # PN[PE.DEMONSTRATOR36]: [["set"] + ("pythonpath=" + pkg_dir.strip()).split(" ") + ['&&', 'python', '-m', 'HSTB.gui.demo'], "", "Pydro367"],
    # PN[PE.DEMONSTRATOR36]: [["pythonpath=%s" % pkg_dir, '&&', 'python', '-m', 'HSTB.gui.demo'], "set", "Pydro367"],
    PN[PE.DEMONSTRATOR27]: [["-m", r"HSTB.gui.demo"], RTE.PYTHON, "Pydro27"],
    PN[PE.WXDEMO27]: [["-m", r"wxPython_demo.demo"], RTE.PYTHON, "Pydro27"],
    PN[PE.WXDEMO36]: [["-m", r"wxPython_demo.demo"], RTE.PYTHON, "Pydro367"],
    PN[PE.SPYDER27]: [[], "spyder", "Pydro27", "", True],
    # PN[PE.PYTHONWIN27]: [[], PathToSitePkgs + '\\pythonwin\\Pythonwin.exe', "Pydro27", ],
    PN[PE.PYTHONWIN27]: [["Pydro27"], path_to_NOAA_site_packages("run_pythonwin.bat"), "base", "", True],
    PN[PE.CONSOLE27]: [[], "", "Pydro27", "Python2", True, True],
    PN[PE.IPYTHON27]: [["--ipython-dir=%s" % docs_path], "ipython.exe", "Pydro27", "", True, True],
    PN[PE.IPYTHONWX27]: [["--pylab=wx", "--ipython-dir=%s" % docs_path], "ipython.exe", "Pydro27", "", True, True],
    PN[PE.IPYTHONQT27]: [["--pylab=qt", "--ipython-dir=%s" % docs_path], "ipython.exe", "Pydro27", "", True, True],
    PN[PE.IPYTHONNOTEBOOK27]: [["notebook", "--notebook-dir=%s" % jupyter_docs], "jupyter", "Pydro27", "", True, True],
    PN[PE.IMAGE_RENAME]: [["-m", "HSTB.gui.renaming_images", ], RTE.PYTHON, "Pydro367"],
    PN[PE.NBS_EMAIL]: [["-m", "HSTB.gui.nbs_transmit", ], RTE.PYTHON, "Pydro367"],
    PN[PE.PROD_EMAIL]: [["-m", "HSTB.gui.product_transmit", ], RTE.PYTHON, "Pydro367"],
    PN[PE.SHPO_EMAIL]: [["-m", "HSTB.gui.shpo_email", ], RTE.PYTHON, "Pydro367"],
    PN[PE.DIR_SIZES]: [["folder_sizes.py", ], RTE.PYTHON, "Pydro367", "Python3\\HSTB\\scripts"],
    PN[PE.SEPERATE_2040_710_FREQ]: [["allfreq.py", ], RTE.PYTHON, "Pydro367", "Python3\\HSTB\\scripts", True, True],
    PN[PE.ENCPRODSPEC]: [["ChangeENCProductSpec.py", ], RTE.PYTHON, "Pydro27", "Python2\\HSTB\\scripts", ],
    PN[PE.MAKECATALOG]: [["-m", "HSTB.gui.make_000_catalog", ], RTE.PYTHON, "Pydro27", "", ],
    PN[PE.PHBCOPY]: [["-m", "HSTB.gui.copy_backscatter", ], RTE.PYTHON, "Pydro27", "", ],
    PN[PE.NCEICHECK]: [["-m", "HSTB.gui.CheckoutNCEI", ], RTE.PYTHON, "Pydro27", "", True],
    PN[PE.GRIDCOMP]: [["-m", "HSTB.gui.surface_comparison", ], RTE.PYTHON, "Pydro27", "", True],
    PN[PE.FETCHTIDES]: [["-m" "HSTB.gui.fetchtides", ], RTE.PYTHON, "Pydro27", "", ],
    PN[PE.CSARQA]: [["-m", r"HSTB.gui.FinalizedCSARsurfaceQA"], RTE.PYTHON, "Pydro27", "", ],
    PN[PE.BDB_ASCII]: [["-m", r"HSTB.gui.BDBExportToASCIIstats"], RTE.PYTHON, "Pydro27", "", ],
    PN[PE.VDATUM_SEP]: [["-m", r"HSTB.gui.VDatumGridFromShapefilePoly"], RTE.PYTHON, "Pydro27", "", ],
    PN[PE.AUTOQC]: [["-m", r"HSTB.gui.POSPacAutoQC"], RTE.PYTHON, "Pydro27", "", ],
    PN[PE.LICENSES27]: [["-m", r"HSTB.gui.licenses", ], RTE.PYTHON, "Pydro27"],
    PN[PE.LICENSES]: [[r"license_gui.py", ], RTE.PYTHON, "Pydro367", "Python3\\HSTB\\gui\\licenses"],
    PN[PE.PERMISSIONS]: [[], "fix_permissions.bat", "", "", True],
    PN[PE.SURVEY_OUTLINES]: [["-m", "HSTB.survey_outline.gui"], RTE.PYTHON, "Pydro367", ""],
    PN[PE.VELOCIPY]: [["-m", r"HSTB.gui.soundspeed"], RTE.PYTHON, "Pydro27"],
    PN[PE.SIMPLE_TCARI]: [["-m", r"HSTB.gui.TCARI", "-p", "0"], RTE.PYTHON, "Pydro27"],
    PN[PE.SIMPLE_TIDES_REQ]: [["-m", r"HSTB.gui.TCARI", "-p", "1"], RTE.PYTHON, "Pydro27"],
    PN[PE.AUVDEPTH]: [["-m", r"HSTB.gui.AUVDepth"], RTE.PYTHON, "Pydro27"],

    PN[PE.VR_BAG]: [["VR_to_SR_Bag.py", ], RTE.PYTHON, "Pydro367", "Python3\\HSTB\\scripts", ],

    PN[PE.BAGEXPLORER]: [["-m", r"hyo2.bagexplorer"], RTE.PYTHON, "Pydro367"],
    PN[PE.BRESS]: [["-m", r"hyo2.bress.app"], RTE.PYTHON, "Pydro367"],
    PN[PE.CATOOLS]: [["-m", r"hyo2.ca.catools"], RTE.PYTHON, "Pydro367"],
    PN[PE.ENCX]: [["-m", r"hyo2.encx"], RTE.PYTHON, "Pydro367"],
    PN[PE.FIGLEAF]: [["-m", r"hyo2.figleaf.app"], RTE.PYTHON, "Pydro367"],
    PN[PE.OPENBST]: [["-m", r"hyo2.openbst.app"], RTE.PYTHON, "Pydro367"],
    PN[PE.QCTOOLS]: [["-m", r"hyo2.qc.qctools"], RTE.PYTHON, "Pydro367"],
    PN[PE.QAX]: [["-m", r"hyo2.qax.app"], RTE.PYTHON, "Pydro367"],
    PN[PE.NOAA_S57]: [["-m", r"hyo2.abc.app.dialogs.noaa_s57"], RTE.PYTHON, "Pydro367"],
    PN[PE.SCRIPT_FLIERS]: [["run_find_fliers_v8.py", ], RTE.PYTHON, "Pydro367",
                           "Python3\\hyo2\\qc\\scripts", ],
    PN[PE.SCRIPT_UNCERTAINTY]: [["run_bag_uncertainty_check.py", ], RTE.PYTHON, "Pydro367",
                                "Python3\\hyo2\\qc\\scripts", ],
    PN[PE.SIS4]: [["run.py", ], RTE.PYTHON, "Pydro367", "Python3\\hyo2\\kng\\emu\\sis4", ],
    PN[PE.SIS5]: [["run.py", ], RTE.PYTHON, "Pydro367", "Python3\\hyo2\\kng\\emu\\kctrl", ],
    PN[PE.SOUNDSPEED]: [["-m", r"hyo2.soundspeedmanager"], RTE.PYTHON, "Pydro367"],
    PN[PE.HDF_COMPASS]: [["-m", r"hdf_compass.compass_viewer"], RTE.PYTHON, "Pydro367"],
    PN[PE.STORMFIX]: [["-m", r"hyo2.stormfix.app"], RTE.PYTHON, "Pydro367"],

    PN[PE.PYTHON_BASICS]: [["notebook", "Python3\\hyo2\\notebooks\\python_basics\\index.ipynb"], "jupyter",
                           "Pydro367", "", True, True],
    PN[PE.REVERT_PB_NOTEBOOKS]: [["Python3\\hyo2\\notebooks\\python_basics"],
                                 path_to_NOAA_site_packages("remove_and_revert.bat"), "", ""],
    PN[PE.OCEAN_DATA_SCIENCE]: [["notebook", "Python3\\hyo2\\notebooks\\ocean_data_science\\index.ipynb"], "jupyter",
                                "Pydro367", "", True, True],
    PN[PE.REVERT_ODS_NOTEBOOKS]: [["Python3\\hyo2\\notebooks\\ocean_data_science"],
                                  path_to_NOAA_site_packages("remove_and_revert.bat"), "", ""],
    PN[PE.PICKY]: [["-m", r"HSTB.picky"], RTE.PYTHON, "Pydro367"],
    PN[PE.SUSSIE]: [["-m", r"oshydro.sussie.app"], RTE.PYTHON, "Pydro367"],
}

ProgramIcons = {
    PN[PE.CASTTIME]: PathToResource("Pydro.ico"),
    PN[PE.CHARLENE]: PathToResource('charlene_AK2_icon.ico'),
    PN[PE.S57COMPARE]: PathToResource("Pydro.ico"),
    PN[PE.ACQFILETRANSFER]: PathToResource("Pydro.ico"),
    PN[PE.SHAM]: PathToResource('charlene_AK2_icon.ico'),
    PN[PE.SCRIBBLE]: PathToResource('charlene_AK2_icon.ico'),
    PN[PE.AUTOQC]: PathToResource("Pydro.ico"),
    PN[PE.CSARQA]: PathToResource("Pydro.ico"),
    PN[PE.BDB_ASCII]: PathToResource("Pydro.ico"),
    PN[PE.SATMON]: PathToResource("Pydro.ico"),
    PN[PE.ROOMBA]: PathToResource("Pydro.ico"),
    PN[PE.VDATUM_SEP]: PathToResource("Pydro.ico"),
    PN[PE.PYDROGIS]: PathToResource("Pydro.ico"),
    PN[PE.POSTACQ]: PathToResource("Pydro.ico"),
    PN[PE.IDLE]: PathToResource("Pydro.ico"),
    PN[PE.FETCHTIDES]: PathToResource("fetchtides.ico"),
    PN[PE.S7K]: PathToResource("Pydro.ico"),
    PN[PE.BENCHMARK]: PathToResource("Pydro.ico"),
    PN[PE.PHBCOPY]: PathToResource("Pydro.ico"),
    PN[PE.MAKECATALOG]: PathToResource("Pydro.ico"),
    PN[PE.ENCPRODSPEC]: PathToResource("Pydro.ico"),
    PN[PE.IMAGE_RENAME]: PathToResource("Pydro.ico"),
    PN[PE.NBS_EMAIL]: PathToResource("branch_dm_tools.ico"),
    PN[PE.PROD_EMAIL]: PathToResource("branch_dm_tools.ico"),
    PN[PE.SHPO_EMAIL]: PathToResource("branch_dm_tools.ico"),
    PN[PE.DIR_SIZES]: PathToResource("Pydro.ico"),
    PN[PE.SEPERATE_2040_710_FREQ]: PathToResource("Pydro.ico"),
    PN[PE.WEEKLYREP]: PathToResource("Pydro.ico"),
    PN[PE.HYPACKLINES]: PathToResource("Pydro.ico"),
    PN[PE.TOGGLE]: PathToResource("Pydro.ico"),
    PN[PE.GRIDCOMP]: PathToResource("Pydro.ico"),
    PN[PE.NCEICHECK]: PathToResource("Pydro.ico"),
    PN[PE.TJ_ACQ_LOG]: PathToResource("Pydro.ico"),

    PN[PE.LTD]: PathToResource("Pydro.ico"),
    PN[PE.PYTHONWIN]: PathToResource("Pydro.ico"),
    PN[PE.CONSOLE367]: PathToResource("Pydro.ico"),
    PN[PE.IPYTHON]: PathToResource("Pydro.ico"),
    PN[PE.IPYTHONWX]: PathToResource("Pydro.ico"),
    PN[PE.IPYTHONQT]: PathToResource("Pydro.ico"),
    PN[PE.IPYTHONNOTEBOOK]: PathToResource("Pydro.ico"),
    PN[PE.SPYDER]: PathToResource("Pydro.ico"),
    PN[PE.SPYDER27]: PathToResource("Pydro.ico"),
    PN[PE.SPYDER38]: PathToResource("Pydro.ico"),
    PN[PE.CREATE38ENV]: PathToResource("Pydro.ico"),
    PN[PE.CONSOLE38]: PathToResource("Pydro.ico"),
    PN[PE.PYTHONWIN27]: PathToResource("Pydro.ico"),
    PN[PE.CONSOLE27]: PathToResource("Pydro.ico"),
    PN[PE.IPYTHON27]: PathToResource("Pydro.ico"),
    PN[PE.IPYTHONWX27]: PathToResource("Pydro.ico"),
    PN[PE.IPYTHONQT27]: PathToResource("Pydro.ico"),
    PN[PE.IPYTHONNOTEBOOK27]: PathToResource("Pydro.ico"),

    PN[PE.VELOCIPY]: PathToResource("Pydro.ico"),
    PN[PE.LICENSES27]: PathToResource("Pydro.ico"),
    PN[PE.LICENSES]: PathToResource("Pydro.ico"),
    PN[PE.PERMISSIONS]: PathToResource("Pydro.ico"),
    PN[PE.SURVEY_OUTLINES]: PathToResource("Pydro.ico"),
    PN[PE.DEMONSTRATOR27]: PathToResource("Pydro.ico"),
    PN[PE.DEMONSTRATOR36]: PathToResource("Pydro.ico"),
    PN[PE.WXDEMO27]: PathToResource("Pydro.ico"),
    PN[PE.WXDEMO36]: PathToResource("Pydro.ico"),
    PN[PE.SIMPLE_TCARI]: PathToResource("Pydro.ico"),
    PN[PE.SIMPLE_TIDES_REQ]: PathToResource("Pydro.ico"),
    PN[PE.AUVDEPTH]: PathToResource("Pydro.ico"),
    PN[PE.XMLDR]: PathToResource("Pydro.ico"),

    PN[PE.SOUNDSPEED]: os.path.join(PathTo_hyo2, r"soundspeedmanager\media\SoundSpeedManager.ico"),
    PN[PE.HDF_COMPASS]: PathToResource("Pydro.ico"),
    PN[PE.SIS4]: PathToResource("Pydro.ico"),
    PN[PE.SIS5]: PathToResource("Pydro.ico"),
    PN[PE.STORMFIX]: os.path.join(PathTo_hyo2, r"stormfix\app\media\StormFix.ico"),
    PN[PE.QCTOOLS]: os.path.join(PathTo_hyo2, r"qc\qctools\media\QCTools.ico"),
    PN[PE.CATOOLS]: os.path.join(PathTo_hyo2, r"ca\catools\media\CATools.ico"),
    PN[PE.QAX]: os.path.join(PathTo_hyo2, r"qax\app\media\QAX.ico"),
    PN[PE.ENCX]: os.path.join(PathTo_hyo2, r"encx\media\ENCX.ico"),
    PN[PE.BAGEXPLORER]: os.path.join(PathTo_hyo2, r"bagexplorer\media\BAGExplorer.ico"),
    PN[PE.FIGLEAF]: os.path.join(PathTo_hyo2, r"figleaf\app\media\figleaf.ico"),
    PN[PE.OPENBST]: os.path.join(PathTo_hyo2, r"openbst\app\media\openbst.ico"),
    PN[PE.BRESS]: os.path.join(PathTo_hyo2, r"bress\app\media\Bress.ico"),
    PN[PE.NOAA_S57]: PathToResource("Pydro.ico"),

    PN[PE.SCRIPT_FLIERS]: PathToResource("Pydro.ico"),
    PN[PE.SCRIPT_UNCERTAINTY]: PathToResource("Pydro.ico"),
    PN[PE.VR_BAG]: PathToResource("Pydro.ico"),

    PN[PE.PYTHON_BASICS]: os.path.join(PathTo_hyo2, r"notebooks\python_basics\images\python_basics.ico"),
    PN[PE.REVERT_PB_NOTEBOOKS]: PathToResource("Pydro.ico"),
    PN[PE.OCEAN_DATA_SCIENCE]: os.path.join(PathTo_hyo2, r"notebooks\ocean_data_science\images\python_basics.ico"),
    PN[PE.REVERT_ODS_NOTEBOOKS]: PathToResource("Pydro.ico"),
    PN[PE.PICKY]: PathToResource("Pydro.ico"),

    PN[PE.SUSSIE]: os.path.join(PathTo_hyo2, r"sussie\app\media\Sussie.ico"),
}

ProgramTreeIcons = {
    PN[PE.SCRIBBLE]: PathToResource("recent.png"),
    PN[PE.NOAA_S57]: PathToResource("recent.png"),
    PN[PE.BRESS]: PathToResource("recent.png"),
    PN[PE.FIGLEAF]: PathToResource("recent.png"),
    PN[PE.SURVEY_OUTLINES]: PathToResource("recent.png"),
    PN[PE.GRIDCOMP]: PathToResource("recent.png"),
    PN[PE.NCEICHECK]: PathToResource("recent.png"),
    PN[PE.SCRIPT_FLIERS]: PathToResource("recent.png"),
    PN[PE.SCRIPT_UNCERTAINTY]: PathToResource("recent.png"),
    PN[PE.VR_BAG]: PathToResource("recent.png"),
    PN[PE.TJ_ACQ_LOG]: PathToResource("recent.png"),

    PN[PE.PYTHON_BASICS]: PathToResource("recent.png"),
    PN[PE.OCEAN_DATA_SCIENCE]: PathToResource("recent.png"),
}
IconNumbers = {}

# Paths to html document files to be displayed when the application is selected in the explorer tree
PYTHON_DOCS = path_to_html("Pydro", "Python.html")
ProgramDocs = {
    "General": path_to_html("Pydro", "General.html"),
    "ESRI-Arc": path_to_html("Arc", "connecting.html"),
    PN[PE.LTD]: path_to_html("Pydro", "LTD.html"),
    PN[PE.CASTTIME]: path_to_html("Apps", "CastTime.html"),
    PN[PE.CHARLENE]: path_to_html("Charlene_Docs", "index.html"),
    PN[PE.S57COMPARE]: path_to_html("Pydro", "General.html"),
    PN[PE.ACQFILETRANSFER]: path_to_html("Pydro", "General.html"),
    PN[PE.SHAM]: path_to_html("Apps", "sham.html"),
    PN[PE.SCRIBBLE]: path_to_html("Apps", "scribble.html"),
    PN[PE.S7K]: path_to_html("Apps", "7kToS7k.html"),
    PN[PE.ENCPRODSPEC]: path_to_html("Apps", "ENCProdSpec.html"),
    PN[PE.IMAGE_RENAME]: path_to_html("Apps", "RenameFFFImages.html"),
    PN[PE.NBS_EMAIL]: path_to_html("Pydro", "General.html"),
    PN[PE.PROD_EMAIL]: path_to_html("Pydro", "General.html"),
    PN[PE.SHPO_EMAIL]: path_to_html("Pydro", "General.html"),
    PN[PE.DIR_SIZES]: path_to_html("Pydro", "General.html"),
    PN[PE.SEPERATE_2040_710_FREQ]: path_to_html("Apps", "SeperateEM_Freq.html"),
    PN[PE.PHBCOPY]: path_to_html("Apps", "PHBQuickTransfer.html"),
    PN[PE.MAKECATALOG]: path_to_html("Apps", "Make000Catalog.html"),
    PN[PE.AUTOQC]: path_to_html("Apps", "PosPacAutoQC.html"),
    PN[PE.CSARQA]: path_to_html("Apps", "FinalizedCSAR_QA.html"),
    PN[PE.BDB_ASCII]: path_to_html("Apps", "BDBSurfaceASCIIExportStats.html"),
    # PN[PE.SATMON]: PathToHSTPDir+"Satmon\\Saturation Monitor.htm"),
    PN[PE.BRESS]: path_to_html("Apps", "Bress.html"),
    PN[PE.SOUNDSPEED]: path_to_html("Apps", "SoundSpeed.html"),
    PN[PE.HDF_COMPASS]: path_to_html("Apps", "hdf_compass.html"),
    PN[PE.STORMFIX]: path_to_html("Apps", "StormFix.html"),
    PN[PE.SIS4]: path_to_html("Apps", "SIS.html"),
    PN[PE.SIS5]: path_to_html("Apps", "SIS.html"),
    PN[PE.QCTOOLS]: path_to_html("Apps", "QCTools.html"),
    PN[PE.NOAA_S57]: path_to_html("Apps", "NOAA_S57.html"),
    PN[PE.CATOOLS]: path_to_html("Apps", "CATools.html"),
    PN[PE.QAX]: path_to_html("Apps", "QAX.html"),
    PN[PE.BAGEXPLORER]: path_to_html("Apps", "BAGExplorer.html"),
    PN[PE.OPENBST]: path_to_html("Apps", "OpenBST.html"),
    PN[PE.FIGLEAF]: path_to_html("Apps", "FigLeaf.html"),
    PN[PE.VDATUM_SEP]: path_to_html("Apps", "VDatumSEPfromShapefile.html"),
    PN[PE.VELOCIPY]: path_to_html("Apps", "velocipy.html"),
    PN[PE.LICENSES27]: path_to_html("Pydro", "licenses.html"),
    PN[PE.LICENSES]: path_to_html("Pydro", "licenses.html"),
    PN[PE.AUVDEPTH]: path_to_html("Apps", "AUVdepth.html"),
    PN[PE.PYDROGIS]: path_to_html("Pydro", "Pydro.html"),
    PN[PE.XMLDR]: path_to_html("Apps", "XmlDR.html"),
    PN[PE.POSTACQ]: path_to_html("Apps", "PostAcqTools.html"),
    PN[PE.SPYDER]: PYTHON_DOCS,
    PN[PE.SPYDER27]: PYTHON_DOCS,
    PN[PE.SPYDER38]: PYTHON_DOCS,
    PN[PE.PYTHONWIN]: PYTHON_DOCS,
    PN[PE.PYTHONWIN27]: PYTHON_DOCS,
    PN[PE.IDLE]: PYTHON_DOCS,
    PN[PE.FETCHTIDES]: path_to_html("Apps", "Fetchtides.html"),
    PN[PE.CREATE38ENV]: PYTHON_DOCS,
    PN[PE.CONSOLE38]: PYTHON_DOCS,
    PN[PE.CONSOLE367]: PYTHON_DOCS,
    PN[PE.CONSOLE27]: PYTHON_DOCS,
    PN[PE.IPYTHON27]: PYTHON_DOCS,
    PN[PE.IPYTHONWX27]: PYTHON_DOCS,
    PN[PE.IPYTHONQT27]: PYTHON_DOCS,
    PN[PE.IPYTHON]: PYTHON_DOCS,
    PN[PE.IPYTHONWX]: PYTHON_DOCS,
    PN[PE.IPYTHONQT]: PYTHON_DOCS,
    PN[PE.WXDEMO27]: PYTHON_DOCS,
    PN[PE.WXDEMO36]: PYTHON_DOCS,
    PN[PE.IPYTHONNOTEBOOK]: PYTHON_DOCS,
    PN[PE.BENCHMARK]: path_to_html("Apps", "Caris_Benchmark_Instructions.html"),
    PN[PE.WEEKLYREP]: path_to_html("Apps", "weekly_reports.html"),
    PN[PE.HYPACKLINES]: path_to_html("Arc", "line_plan_tools.html"),
    PN[PE.PERMISSIONS]: path_to_html("Pydro", "FixPermissions.html"),
    PN[PE.SURVEY_OUTLINES]: path_to_html("Pydro", "ExtractSurveyOutlines.html"),
    PN[PE.TOGGLE]: path_to_html("Apps", "Toggle.html"),
    PN[PE.DEMONSTRATOR27]: path_to_html("Apps", "CodeBaseDemo.html"),
    PN[PE.DEMONSTRATOR36]: path_to_html("Apps", "CodeBaseDemo.html"),
    PN[PE.ENCX]: path_to_html("Apps", "ENCX.html"),
    PN[PE.SIMPLE_TCARI]: path_to_html("Apps", "SimpleTCARI.html"),
    PN[PE.SIMPLE_TIDES_REQ]: path_to_html("Apps", "SimpleTidesRequest.html"),
    PN[PE.GRIDCOMP]: path_to_html("Apps", "CSAR_Surface_Comparison.html"),
    PN[PE.NCEICHECK]: path_to_html("Apps", "NCEI_Checkout_Tool.html"),
    PN[PE.SCRIPT_FLIERS]: path_to_html("Apps", "script_fliers.html"),
    PN[PE.SCRIPT_UNCERTAINTY]: path_to_html("Apps", "script_uncertainty.html"),
    PN[PE.VR_BAG]: path_to_html("Apps", "vr_to_sr_bag.html"),

    PN[PE.PYTHON_BASICS]: path_to_html("Apps", "python_basics.html"),
    PN[PE.REVERT_PB_NOTEBOOKS]: path_to_html("Apps", "python_basics.html"),
    PN[PE.OCEAN_DATA_SCIENCE]: path_to_html("Apps", "ocean_data_science.html"),
    PN[PE.REVERT_ODS_NOTEBOOKS]: path_to_html("Apps", "ocean_data_science.html"),
    PN[PE.PICKY]: path_to_html("Apps", "Picky.html"),
    PN[PE.SUSSIE]: path_to_html("Apps", "Sussie.html"),
}

# this description can be shown in a tooltip or list of programs in the docs.
program_simple_descr = {
    PN[PE.CASTTIME]: """Monitoring and evaluation tool to determine how often to take sound speed profiles""",
    PN[PE.CHARLENE]: """Automated processing and data transfer tool""",
    PN[PE.S57COMPARE]: """Compare two s57 files to get field/branch differences""",
    PN[PE.ACQFILETRANSFER]: """Automated Launch Transfer Drive and Directory Monitoring Tool""",
    PN[PE.SHAM]: """Shoreline attribution tool""",
    PN[PE.SCRIBBLE]: """Check project structure, generate miles report, auto populate your XMLDR""",
    PN[PE.LTD]: """Create a Letter Transmitting Data (LTD) NOAA form 61-29 as a PDF document""",
    PN[PE.S7K]: """Converts a Hypack 7K file to a Reson s7k file and adds the navigation and attitude stored in the corresponding Hypack HSX file into the s7k""",
    PN[PE.ENCPRODSPEC]: """Changes the ENC Product Spec in the selected .000 file from a value of 31 to 1""",
    PN[PE.IMAGE_RENAME]: """Rename FFF Images based on HTD 2018-4""",
    PN[PE.NBS_EMAIL]: """NBS Mass Transmittal Email""",
    PN[PE.PROD_EMAIL]: """HSD Digital Production Transmittal Email""",
    PN[PE.SHPO_EMAIL]: """SHPO Email""",
    PN[PE.DIR_SIZES]: """Check directory sizes and report in a CSV format""",
    PN[PE.SEPERATE_2040_710_FREQ]: """Seperate EM2040 and EM710 by Frequency""",
    PN[PE.PHBCOPY]: """Automates both directory creation and the transfer of backscatter component files""",
    PN[PE.MAKECATALOG]: """Automatically create a Catalog.031 for a directory structure""",
    PN[PE.AUTOQC]: """A mechanization of the SBET Solution Quality Assessment guidelines as documented in the POSPac MMS GNSS-Inertial Tools User Guide""",
    PN[PE.CSARQA]: """Computes basic statistics to assess finalized surface sounding density and uncertainty compliance per the NOS Hydrographic Surveys Specifications and Deliverables""",
    PN[PE.BDB_ASCII]: """produces a statistical summary of a scalar field dataset""",
    PN[PE.BRESS]: """A library and an app for landform identification and seafloor segmentation""",
    PN[PE.SOUNDSPEED]: """Advanced sound speed management tool - reads, processes, reformats and exports sound speed casts""",
    PN[PE.HDF_COMPASS]: """View HDF5 datasets, attributes, and groups. Simple line, image, and contour plots are supported as well""",
    PN[PE.STORMFIX]: """Provides means to identify and reduce the presence of artifacts in a backscatter mosaic""",
    PN[PE.SIS4]: """A simple application to emulate Kongsberg SIS 4 interaction """,
    PN[PE.SIS5]: """A simple application to emulate Kongsberg K-Ctrl interaction """,
    PN[PE.QCTOOLS]: """Provides tools for quality control and review of survey data""",
    PN[PE.CATOOLS]: """Provides tools for verifying the adequacy of nautical charts""",
    PN[PE.QAX]: """Provides tools for quality assurance of ocean mapping data""",
    PN[PE.NOAA_S57]: """Installation of NOAA S57 Support Files for CARIS software""",
    PN[PE.BAGEXPLORER]: """Browse and interact with Bathymetric Attributed Grid (BAG) files""",
    PN[PE.OPENBST]: """A library and an app for processing acoustic backscatter data""",
    PN[PE.FIGLEAF]: """A library and an app for manipulation of raster data""",
    PN[PE.VDATUM_SEP]: """Computes a datum height (aka ellipsoid separation or SEP) grid""",
    PN[PE.VELOCIPY]: """(Deprecated for SoundSpeedManager) Python version of NOAA's Velocity programs - reads, processes, reformats and exports sound speed casts""",
    PN[PE.LICENSES27]: """Shows the licenses of software distributed within Pydro""",
    PN[PE.LICENSES]: """Shows the licenses of software distributed within Pydro""",
    PN[PE.PERMISSIONS]: """Fixes permissions of files that erroneously got admin only (will prompt for admin to fix)""",
    PN[PE.SURVEY_OUTLINES]: """Extracts Survey Outlines from Bag/Rasters to GeoPackages (GIS files)""",
    PN[PE.AUVDEPTH]: """AUV pressure to depth""",
    PN[PE.PYDROGIS]: """Specialized GIS which houses multiple tools/abilities such as creating TCARI files, generating reports and managing survey features""",
    PN[PE.XMLDR]: """Data entry application to create various reports, stored in PDF and printable to PDF""",
    PN[PE.POSTACQ]: """Set of tools is aimed at fixing data issues in raw data""",
    PN[PE.TJ_ACQ_LOG]: """ Program that will take the positioning from SIS, time from the computer, and any text notes written and create a geopackage that can be opened in Caris 11 """,
    PN[PE.SPYDER]: """Python IDE""",
    PN[PE.SPYDER27]: """Python IDE""",
    PN[PE.SPYDER38]: """Python IDE""",
    PN[PE.PYTHONWIN]: """Python IDE""",
    PN[PE.PYTHONWIN27]: """Python IDE""",
    PN[PE.IDLE]: """Python IDE""",
    PN[PE.FETCHTIDES]: """Tool for downloading, storing and exporting tide data from NOAA's |COOPS|""",
    PN[PE.CONSOLE27]: """Python IDE""",
    PN[PE.CONSOLE367]: """Python IDE""",
    PN[PE.CREATE38ENV]: """Python IDE""",
    PN[PE.CONSOLE38]: """Python IDE""",
    PN[PE.IPYTHON27]: """Python IDE""",
    PN[PE.IPYTHONWX27]: """Python IDE""",
    PN[PE.IPYTHONQT27]: """Python IDE""",
    PN[PE.IPYTHON]: """Python IDE""",
    PN[PE.IPYTHONWX]: """Python IDE""",
    PN[PE.IPYTHONQT]: """Python IDE""",
    PN[PE.IPYTHONNOTEBOOK]: """Python IDE""",
    PN[PE.WEEKLYREP]: """Arc + Caris tool to create weekly report tif images""",
    PN[PE.HYPACKLINES]: """Arc tool to create survey lines""",
    PN[PE.DEMONSTRATOR27]: """Development environment to show and test code and modules within the Pydro distribution""",
    PN[PE.DEMONSTRATOR36]: """Development environment to show and test code and modules within the Pydro distribution""",
    PN[PE.WXDEMO27]: """Demostrates WX graphical user interface""",
    PN[PE.WXDEMO36]: """Demostrates WX graphical user interface""",
    PN[PE.ENCX]: """Tools to explore the ENC data content at multiple levels""",
    PN[PE.SIMPLE_TCARI]: """Basic interface to apply TCARI tides data""",
    PN[PE.SIMPLE_TIDES_REQ]: """Tool to request tides from co-ops (for NOAA surveys)""",
    PN[PE.GRIDCOMP]: """Tool to analyze the difference between two gridded Depth/Elevation layers in CSAR/BAG format""",
    PN[PE.NCEICHECK]: """NCEI Checkout Tool allows users to validate the required naming conventions and the bottom sample ascii files in a NCEI submission directory""",
    PN[PE.SCRIPT_FLIERS]: """Script to Find Fliers""",
    PN[PE.SCRIPT_UNCERTAINTY]: """Script for Empty Uncertainty""",
    PN[PE.VR_BAG]: """Tool to convert from VR Surface to SR Bag""",

    PN[PE.PYTHON_BASICS]: """Open the Programming Basics with Python notebooks""",
    PN[PE.REVERT_PB_NOTEBOOKS]: """Return the Programming Basics with Python notebooks to their initial state (removes local changes)""",
    PN[PE.OCEAN_DATA_SCIENCE]: """Open the Introduction to Ocean Data Science notebooks""",
    PN[PE.REVERT_ODS_NOTEBOOKS]: """Return the Introduction to Ocean Data Science notebooks to their initial state (removes local changes)""",
    PN[PE.PICKY]: """Automated Side Scan Detection""",
    PN[PE.SUSSIE]: """A collection of tools providing functionalities to handle hydrographic survey data""",
}

for name in ProgramNames.values():  # cheap conversion of all the original dictionaries to Program objects
    kwargs = {}
    try:
        opts = ProgramOpts[name]
        if not isinstance(opts, ProgOpts):
            opts = ProgOpts(*ProgramOpts[name])
        kwargs['run_opts'] = opts
    except KeyError:
        pass
    try:
        kwargs['desktop_icon'] = ProgramIcons[name]
    except KeyError:
        pass
    try:
        kwargs['descr'] = program_simple_descr[name]
    except KeyError:
        pass
    try:
        kwargs['docs'] = ProgramDocs[name]
    except KeyError:
        pass
    try:
        kwargs['tree_icon'] = ProgramTreeIcons[name]
    except KeyError:
        pass

    np = Program(name, **kwargs)



class MyCT(treemixin.ExpansionState, CT.CustomTreeCtrl):
    pass


class XmlDRFrame(BaseAuiFrame.HSTP_AUI_Frame):
    def MakeRST(self):
        output_name = PathToDocs("..", "Docs_source", "Pydro", "program_list_auto.rst")
        output_file = open(output_name, "wb")
        output_file.write("""
=============================
Programs distributed in Pydro
=============================

        """)
        for group in self._ZfileMenuSection:
            self._AddGroupToRST(group, output_file)

        toc_output_name = PathToDocs("..", "Docs_source", "Apps", "index_all_apps.rst")
        output_file = open(toc_output_name, "wb")
        output_file.write("""
=================================
All Programs distributed in Pydro
=================================
.. toctree::
   :maxdepth: 3

""")
        progs = list(ProgramList.keys())
        progs.sort()
        for p in progs:
            rst_path = ProgramList[p].docs.replace(PathToDocs("html"), PathToDocs("..\\Docs_source")).replace(".html", ".rst")
            if os.path.exists(rst_path):
                entry = ProgramList[p].docs.replace(PathToDocs("html"), "..").replace("\\", "/").replace(".html", "")
                # '../Apps/7kToS7k'
            else:
                # Make a relative link to the html page -- but relative links to non-sphinx docs are not supported currently
                # Found a hack at
                # https://stackoverflow.com/questions/27979803/external-relative-link-in-sphinx-toctree-directive
                entry = ProgramList[p].docs.replace(PathToDocs(), "../..").replace("\\", "/") + "#://"
                # Switch to external link syntax if/when they support it
                # https://github.com/sphinx-doc/sphinx/issues/701
                # https://github.com/sphinx-doc/sphinx/pull/1800
                # output_file.write("   `" + p + " <" + entry + ">`_\n")
                # '../../html/Apps/7kToS7k.html'
            output_file.write("   " + p + " <" + entry + ">\n")

    def _AddGroupToRST(self, group, output_file, headertype="-"):
        groupname, actions = group[:2]
        output_file.write("""
{}
{}
\n""".format(groupname, headertype * len(groupname)))
        if (actions):
            for action in actions[0]:
                if action:
                    if isinstance(action, BaseAuiFrame.HSTPMenuGroup):
                        self._AddGroupToRST(action, output_file, headertype="^")
                    else:
                        name = action[0]
                        try:
                            local_doc_link = ProgramList[name].docs.replace(PathToDocs(), "../..").replace("\\", "/")
                            output_file.write("  - `{} <{}>`_ \n".format(name, local_doc_link))
                            output_file.write("    :: {} \n".format(ProgramList[name].descr))
                        except KeyError:
                            print("{} didn't have a documentation entry".format(name))
                else:
                    output_file.write("\n")
        else:
            output_file.write("\n")

    def __init__(self, parent, id, title):
        G = BaseAuiFrame.HSTPMenuGroup
        I = BaseAuiFrame.HSTPMenuItem
        self._ZfileMenu = G('&File', [])
        self._ZfileMenuSection = [
            G("New", [[
                I(PN[PE.SCRIBBLE], self),
                I(PN[PE.NOAA_S57], self),
                I(PN[PE.BRESS], self),
                I(PN[PE.FIGLEAF], self),
                I(PN[PE.OPENBST], self),
                I(PN[PE.CATOOLS], self),
                I(PN[PE.QAX], self),
                I(PN[PE.STORMFIX], self),
                # I(PN[PE.ROOMBA], self),
                I(PN[PE.GRIDCOMP], self),
                I(PN[PE.NCEICHECK], self),
                I(PN[PE.SCRIPT_FLIERS], self),
                I(PN[PE.SCRIPT_UNCERTAINTY], self),
                I(PN[PE.SIS5], self),
                I(PN[PE.SURVEY_OUTLINES], self),
                I(PN[PE.VR_BAG], self),
                I(PN[PE.LTD], self),
                I(PN[PE.PYTHON_BASICS], self),
                I(PN[PE.OCEAN_DATA_SCIENCE], self),
            ]], -1),
            G("Backscatter", [[
                I(PN[PE.BRESS], self),
                I(PN[PE.OPENBST], self),
                I(PN[PE.SATMON], self),
                I(PN[PE.S7K], self),
                I(PN[PE.STORMFIX], self),
                I(HSRR.name, self),
            ]], -1),
            G('Sound Speed', [[
                I(PN[PE.VELOCIPY], self),
                I(PN[PE.CASTTIME], self),
                I(PN[PE.SOUNDSPEED], self),
                I(PN[PE.HDF_COMPASS], self),
                I(PN[PE.AUVDEPTH], self),
            ]], -1),
            G('Deliverables', [[
                I(PN[PE.SCRIBBLE], self),
                I(PN[PE.XMLDR], self),
                I(PN[PE.BAGEXPLORER], self),
                I(PN[PE.AUTOQC], self),
                I(PN[PE.LTD], self),
                I(PN[PE.IMAGE_RENAME], self),
                I(PN[PE.SUSSIE], self),
                I(HSRR.name, self),
                # I(PN[PE.CSARQA], self), # depricated; use QC Tools instead
            ]], -1),
            G('ERS', [[
                I(PN[PE.AUTOQC], self),
                I(PN[PE.VDATUM_SEP], self),
            ]], -1),
            G('Surfaces', [[
                I(PN[PE.GRIDCOMP], self),
                I(PN[PE.BDB_ASCII], self),
                I(PN[PE.VDATUM_SEP], self),
                I(PN[PE.BRESS], self),
                I(PN[PE.BAGEXPLORER], self),
                I(PN[PE.FIGLEAF], self),
                I(PN[PE.SUSSIE], self),
                # I(PN[PE.ROOMBA], self),
            ]], -1),
            G('Branch Tools', [[
                I(PN[PE.NOAA_S57], self),
                I(PN[PE.QCTOOLS], self),
                I(PN[PE.CATOOLS], self),
                I(PN[PE.QAX], self),
                I(PN[PE.ENCX], self),
                I(PN[PE.PHBCOPY], self),
                I(PN[PE.MAKECATALOG], self),
                I(PN[PE.ENCPRODSPEC], self),
                I(PN[PE.NCEICHECK], self),
                I(PN[PE.SCRIPT_FLIERS], self),
                I(PN[PE.SCRIPT_UNCERTAINTY], self),
                I(PN[PE.SURVEY_OUTLINES], self),
                I(PN[PE.IMAGE_RENAME], self),
                I(PN[PE.NBS_EMAIL], self),
                I(PN[PE.PROD_EMAIL], self),
                I(PN[PE.SHPO_EMAIL], self),
                I(PN[PE.VR_BAG], self),
                I(PN[PE.FIGLEAF], self),
                I(lnm_calc.name, self),
                I(PN[PE.S57COMPARE], self),
                # I(PN[PE.ROOMBA], self),
                I(PN[PE.SUSSIE], self),
            ]], -1),
            G('ESRI-Arc', [[
                I(PN[PE.HYPACKLINES], self),
                I(beets.name, self),
                I(PN[PE.SURVEY_OUTLINES], self),
            ]], -1),
            G('Tides and Datums', [[
                I(PN[PE.PYDROGIS], self),
                I(PN[PE.FETCHTIDES], self),
                I(PN[PE.SIMPLE_TCARI], self),
                I(PN[PE.SIMPLE_TIDES_REQ], self),
                I(PN[PE.VDATUM_SEP], self),
            ]], -1),
            G("Raw Data Access/Conversion", [[
                I(PN[PE.S7K], self),
                I(PN[PE.CHARLENE], self),
                I(PN[PE.ACQFILETRANSFER], self),
                I(PN[PE.SHAM], self),
                I(PN[PE.STORMFIX], self),
            ]], -1),
            G("Learning", [[
                I(PN[PE.DEMONSTRATOR27], self),
                I(PN[PE.DEMONSTRATOR36], self),
                I(PN[PE.PYTHON_BASICS], self),
                I(PN[PE.REVERT_PB_NOTEBOOKS], self),
                I(PN[PE.OCEAN_DATA_SCIENCE], self),
                I(PN[PE.REVERT_ODS_NOTEBOOKS], self),
            ]], -1),
            G("Supplemental Data", [[
                I(download_aviso.name, self),
                I(download_gebco_gdb.name, self),
                I(download_gebco_rasters.name, self),
            ]], -1),
            G('Other', [[
                I(PN[PE.TOGGLE], self),
                I(PN[PE.LICENSES27], self),
                I(PN[PE.LICENSES], self),
                I(PN[PE.PYDROGIS], self),
                I(PN[PE.POSTACQ], self),
                # I(PN[PE.BENCHMARK], self),
                I(PN[PE.PERMISSIONS], self),
                G('Python 3.6 shells and editors', [[
                    I(PN[PE.SPYDER], self),
                    I(PN[PE.PYTHONWIN], self),
                    I(PN[PE.IPYTHON], self),
                    I(PN[PE.IPYTHONNOTEBOOK], self),
                    I(PN[PE.CONSOLE367], self),
                    I(PN[PE.WXDEMO36], self),
                ]], -1),
                G('Python 2.7 shells and editors', [[
                    I(PN[PE.SPYDER27], self),
                    I(PN[PE.PYTHONWIN27], self),
                    I(PN[PE.IPYTHON27], self),
                    I(PN[PE.IPYTHONNOTEBOOK27], self),
                    I(PN[PE.IPYTHONWX27], self),
                    I(PN[PE.CONSOLE27], self),
                    I(PN[PE.WXDEMO27], self),
                ]], -1),
                G('Python 3.8 shells and editors', [[
                    I(PN[PE.CREATE38ENV], self),
                    I(PN[PE.SPYDER38], self),
                    I(PN[PE.CONSOLE38], self),
                ]], -1),
                I(PN[PE.IDLE], self),
                I(PN[PE.AUVDEPTH], self),
                I(PN[PE.SIS4], self),
                I(PN[PE.SIS5], self),
                I(PN[PE.PICKY], self),
            ]], -1),

            G("BETA / EXPERIMENTAL", [[
                I(PN[PE.DIR_SIZES], self),
                I(PN[PE.WEEKLYREP], self),
                I(PN[PE.SEPERATE_2040_710_FREQ], self),
                I(PN[PE.ROOMBA], self),
                I(PN[PE.TJ_ACQ_LOG], self),
                I(files_checker.name, self),
            ]], -1),
        ]
#        self._WindowMenu = G('&Window',[])
#        self._WindowMenuSection = [
#                                   I('Save Perspective', self),
#                                   I('Load Perspective', self),
#                                   I('Reload', self),
#                                 ]
        self._InternalEvents = {}  # {'TestInternal':[self, 'Test2',-1]}
        _Zevents = {}  # dictionary to keep events in
        fullmenu = [self._ZfileMenu, ]
        self.DisableMenus = []
        self.pickle_fname = path_to_HSTB("RecentlyRun.pickle")
        self.recent = []
        BaseAuiFrame.HSTP_AUI_Frame.__init__(self, parent, -1, title, "LauncherApp", self._InternalEvents, [], _Zevents, self.DisableMenus, fullmenu)

    def OnPaneClose(self, event):
        # self.docManager.PaneClosing(event)
        event.Skip()

    def OnCloseWindow(self, event):
        BaseAuiFrame.HSTP_AUI_Frame.OnCloseWindow(self, event)

    def CreateArgs(self, run_opts):  # args=[], startProg=RunTypeEnum.PYTHON, env="", start_directory="", persistent_env=False):
        ''' HSTP_directory is the start-in folder relative to the HSTP site-package
        args is a list of additional arguments to be passed to the executable
        startProg is either a RunTypeEnum specifying the which executable to use or a path to an executable (or executable name if it's in the path)
        env is the conda environment to run in, if applicable
        Everything is being returned as a "short path" (meaning no spaces on windows) so that shortcuts work on the start menu with multiple commands

        NOTE when run from the application any arguments with spaces will get quotes around them
        but when run from an icon they will just be joined together with spaces.
        ex:  set pythonpath=test => "set pythonpath=test" in subprocess.popen fails but would have worked from the icon
        using: ["set", "pythonpath=test"] works in both icon and subprocess
        '''

        cmd_switch = "/K" if run_opts.persist_console else "/C"
        if run_opts.cmd == RunTypeEnum.PYTHON or run_opts.cmd == RunTypeEnum.RAW:
            if run_opts.env:  # run in the specified environment
                pathToExe = "python"
            else:  # run in the local/current python interpreter
                pathToExe = PathToSitePkgs[:PathToSitePkgs.lower().find('lib')] + "python.exe"
        else:
            pathToExe = run_opts.cmd
        args = copy.copy(run_opts.args)
        if pathToExe:
            try:
                pathToExe = get_short_path_name(pathToExe)
            except Exception:
                pass  # win32api will raise an "error" if something isn't a full path (like "python" using the system path)
            args = [pathToExe] + args
        if run_opts.env:
            pathToActivate = get_short_path_name(path_to_NOAA("..\\Scripts\\activate.bat"))
            args = create_env_cmd_list(run_opts.env, run_opts.persist_console) + args
        if args[-1].endswith("&&"):
            args[-1] = args[-1][:-2]
        if not args[-1]:  # remove the last element if it's blank (may have been a final "&&")
            args.pop(-1)
        full_start_directory = get_short_path_name(path_to_NOAA_site_packages(run_opts.dir))
        if full_start_directory[-1] in ("\\/"):
            full_start_directory = full_start_directory[:-1]
        # subprocess.call([r'%s'%pathToExe, filename,]+ ['"%s"'%a for a in args])
        sub_args = [full_start_directory] + [str(a) for a in args]
        return sub_args

    def _Launch(self, run_opts):  # args=[], startProg=RunTypeEnum.PYTHON, env="", start_directory="", new_console=False, persistent_env=False):
        if run_opts.dir or run_opts.args or run_opts.cmd not in RunTypeEnum:  # don't run if there aren't arguments (ArcMap tools etc.)
            sub_args = self.CreateArgs(run_opts)
            print(sub_args)
            os.chdir(sub_args[0])
            subprocess.Popen(sub_args[1:], creationflags=subprocess.CREATE_NEW_CONSOLE * bool(run_opts.new_console))

            #        to run as admin -- used the example syntax below
            #        import os,sys
            #        import win32com.shell.shell as shell
            #        shell.ShellExecuteEx(lpVerb='runas', lpFile=sys.executable, lpParameters=cmdline_string)

        os.chdir(path_to_HSTB())

    def Launch(self, programName, evt=None, dbg=False):
        '''evt is a placeholder so that menu items can be set up with functools.partial commands to run directly form the menus.
        '''
        self.recent.append(programName)
        self.log.write("Launching %s\n" % programName)
        opts = ProgramList[programName].opts
        if dbg:
            opts = opts.copy()
            opts.persist_console = True  # force the program to have it's own console and remain after closing (the /K option for cmd.exe)
            opts.new_console = True
        self._Launch(opts)
        pickle.dump(self.recent, open(self.pickle_fname, "w+"))

    def OnCreateDesktopIcon(self, evt):
        item = self.tree.GetSelection()
        self.CreateIcon(self.tree.GetItemText(item), "Desktop")

    def OnCreateStartMenuIcon(self, evt):
        item = self.tree.GetSelection()
        self.CreateIcon(self.tree.GetItemText(item), "Programs")

    def CreateIcon(self, prog, place="StartMenu"):
        ''' prog is the program name that matches the keys for the ProgramList dictionary and the menu choices.
        place is a text string that a shell <made from win32com.client.Dispatch("WScript.Shell")> uses to determine the home folder.
          acceptable string values are (https://msdn.microsoft.com/en-us/library/0ea7b5xe%28v=vs.84%29.aspx):
              AllUsersDesktop,  AllUsersStartMenu, AllUsersPrograms,  AllUsersStartup
              Desktop, Programs, StartMenu
              Favorites, Fonts, MyDocuments,    NetHood,    PrintHood,   Recent
              SendTo,  Startup Templates
        '''
        # num_args = len(inspect.getargspec(self.CreateArgs).args) - 1  # fill all args from the datastructure (remove one for implicit self argument)
        args = self.CreateArgs(ProgramList[prog].opts)

        # >>> from win32com.shell import shellcon
        # >>> import win32com.shell.shell
        # >>> win32com.shell.shell.SHGetSpecialFolderPath(0,shellcon.CSIDL_COMMON_STARTMENU)
        # u'C:\\ProgramData\\Microsoft\\Windows\\Start Menu'
        # >>> win32com.shell.shell.SHGetSpecialFolderPath(0,shellcon.CSIDL_STARTMENU)
        # u'C:\\Users\\Barry.Gallagher\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu'

        # >>> import win32com.client
        # >>> objShell = win32com.client.Dispatch("WScript.Shell")
        # >>> allUserProgramsMenu = objShell.SpecialFolders("AllUsersPrograms")
        # >>> objShell.SpecialFolders("AllUsersPrograms")
        # u'C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs'
        # >>> objShell.SpecialFolders("StartMenu")
        # u'C:\\Users\\Barry.Gallagher\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu'
        # >>> shell.SHGetSpecialFolderPath(0,shellcon.CSIDL_DESKTOP)
        # u'C:\\Users\\Barry.Gallagher\\Desktop'
        # >>> win32com.shell.shell.IsUserAnAdmin()
        # False
        shell = Dispatch('WScript.Shell')
        root_path = shell.SpecialFolders(place)
        version_type = (" " + Constants.PydroVersionType()) if Constants.PydroVersionIsDev() else ""
        if place in ("Desktop", "AllUsersDesktop"):
            path = root_path + "\\%s%s.lnk" % (prog, version_type)  # no folder containing icons.
        else:
            path = root_path + "\\PydroXL_19%s\\%s%s.lnk" % (version_type, prog, version_type)
        if (not os.path.isdir(os.path.split(path)[0])):
            os.makedirs(os.path.split(path)[0])  # create the program group
        shortcut = shell.CreateShortCut(path)
        if ProgramList[prog].opts.cmd == RunTypeEnum.PYTHON:
            ind = 3 if args[1].lower() == "cmd.exe" and args[2][:2] in ("/C", "/K") else 1
            args.insert(ind, "&")
            args.insert(ind, get_short_path_name(os.path.join(noaa_sitepkg_dir, "..", "get_updates.bat")))
        # replace double & with single for difference between launching as a string and launching multiple programs in a shortcut
        for i, a in enumerate(args):
            if a.endswith("&&"):
                args[i] = a[:-1]
        shortcut.Targetpath = args[1]
        if len(args) > 2:
            shortcut.Arguments = ' '.join(args[2:])
            # shortcut.Arguments=' "'+'" "'.join(args[2:])+'"'
        shortcut.WorkingDirectory = args[0]
        try:
            icon_path = ProgramList[prog].desktop_icon
            if icon_path is None:
                raise Exception('missing icon spec')
        except:
            icon_path = PathToResource("Pydro.ico")

        shortcut.IconLocation = icon_path
        shortcut.save()

    def OnDebugProgram(self, evt):
        item = self.tree.GetSelection()
        self.Launch(self.tree.GetItemText(item), dbg=True)

    def OnRunProgram(self, evt):
        item = self.tree.GetSelection()
        self.Launch(self.tree.GetItemText(item))

    # def OnPydroGIS(self, event):
    #     self.Launch('PydroGIS')

    def MakeMenuList(self):
        # should make this generic by iterating a list of data/view modules that have menus to display.
        # G = BaseAuiFrame.HSTPMenuGroup

        self._ZfileMenu.RemoveSubItems()
        self._ZfileMenu.AppendSection(self._ZfileMenuSection)

#        self._WindowMenu.RemoveSubItems()
#        self._WindowMenu.AppendSection(self._WindowMenuSection)

        fullmenu = [self._ZfileMenu, ]
        return fullmenu

    def ReloadMenus(self):
        # rebuild the menus anytime we load/remove a module with menu items
        self.CreateNewMenuBar(self._InternalEvents, self.MakeMenuList(), self.DisableMenus)

    def CreateZFrameLayout(self):
        self.CreateLog()  # Create self.log
        if not _dHSTP:
            sys.stdout = self.log
            sys.stderr = self.log

        self.CreateTree()
        self.CreateHTMLWindow()
        wx.Log.SetActiveTarget(wx.LogTextCtrl(self.log))
        if _dHSTP:
            self.CreateShellWindow({'frame': self})
        self._mgr.Update()  # "commit" all changes made to FrameManager
        # now that the various windows with menus are loaded, recreate the main menu
        self.ReloadMenus()

    def CreateTree(self):
        # Set up the contact tree control with log window
        self.tree_panel = wx.Panel(self, style=wx.TAB_TRAVERSAL | wx.CLIP_CHILDREN)
        self.tree = MyCT(self.tree_panel, style=wx.TR_DEFAULT_STYLE | wx.TR_HAS_VARIABLE_ROW_HEIGHT)
        self.tree.SetAGWWindowStyleFlag(CT.TR_HIDE_ROOT | CT.TR_HAS_BUTTONS)
        self._mgr.AddPane(self.tree_panel, wx.aui.AuiPaneInfo().
                          Name("Tree").Caption("Applications").
                          Left().Layer(1).CloseButton(False).MaximizeButton(True).BestSize(wx.Size(300, 600)).FloatingSize((400, 400)))
        if 0:
            self.tree.Bind(wx.EVT_TREE_ITEM_EXPANDED, self.OnItemExpanded)
            self.tree.Bind(wx.EVT_TREE_ITEM_COLLAPSED, self.OnItemCollapsed)
            self.tree.Bind(wx.EVT_LEFT_DOWN, self.OnTreeLeftDown)
        self.tree.Bind(wx.EVT_TREE_SEL_CHANGED, self.OnSelChanged)
        self.tree.Bind(wx.EVT_LEFT_DCLICK, self.OnLeftDClick)

        self.filter = wx.SearchCtrl(self.tree_panel, style=wx.TE_PROCESS_ENTER)
        self.filter.ShowCancelButton(True)
        self.filter.Bind(wx.EVT_TEXT, self.FillTreeItems)
        self.filter.Bind(wx.EVT_SEARCHCTRL_CANCEL_BTN,
                         lambda e: self.filter.SetValue(''))
        self.filter.Bind(wx.EVT_TEXT_ENTER, self.OnSearch)

        searchMenu = wx.Menu()
        item = searchMenu.AppendRadioItem(-1, "Names")
        self.Bind(wx.EVT_MENU, self.OnSearchMenu, item)
        item = searchMenu.AppendRadioItem(-1, "Names, Docs")
        self.Bind(wx.EVT_MENU, self.OnSearchMenu, item)
        self.filter.SetMenu(searchMenu)
        searchMenu.Check(item.GetId(), True)  # Set the docs to be searched by default

        szr = wx.BoxSizer(wx.VERTICAL)
        szr.Add(self.tree, 1, wx.EXPAND)
        szr.Add(wx.StaticText(self.tree_panel, label="Filter Apps:"), 0, wx.TOP | wx.LEFT, 5)
        szr.Add(self.filter, 0, wx.EXPAND | wx.ALL, 5)
        self.tree_panel.SetSizer(szr)

        self.FillTreeItems()

    def ClearEmptyBranches(self, item=None):
        if not item:
            item = self.root
        for c in self.tree.GetItemChildren(item):
            self.ClearEmptyBranches(c)
        if not self.tree.GetItemChildren(item):
            if item.GetText() not in ProgramList:
                self.tree.Delete(item)
        return

    def OnSearch(self, event):
        self.FillTreeItems()

    def OnSearchMenu(self, event):
        self.FillTreeItems()

    def FillTreeItems(self, event=None):
        self.tree.Freeze()
        self.tree.DeleteAllItems()
        self.root = self.tree.AddRoot("Categories")

        il = wx.ImageList(16, 16)
        for k, p in ProgramList.items():
            v = p.tree_icon
            if v:
                img = wx.Bitmap(v)
                # icon = wx.IconFromBitmap(img.ConvertToBitmap() )
                IconNumbers[k] = il.Add(img)
                self.tree.SetImageList(il)

        searchMenu = self.filter.GetMenu().GetMenuItems()
        user_filter = str(self.filter.GetValue())
        filter_name = user_filter
        filter_docs = ""
        if searchMenu[1].IsChecked():
            filter_docs = user_filter

        # Create Recent Run list
        recent_group = self.CreateRecentItemsList()
        self.AddMenuItemsToTree([recent_group], self.root, filter_name=filter_name, filter_docs=filter_docs)
        self.AddMenuItemsToTree(self._ZfileMenuSection, self.root, filter_name=filter_name, filter_docs=filter_docs)
        self.root.GetChildren()[0].Expand()
        self.root.Expand()
        self.ClearEmptyBranches()
        if filter_name or filter_docs:
            self.tree.ExpandAll()
        self.tree.Thaw()

    def CreateRecentItemsList(self):
        try:
            self.recent = pickle.load(open(self.pickle_fname, "r"))
            self.recent = self.recent[-40:]
        except:  # no file
            pass
        c = collections.Counter(self.recent)
        group = BaseAuiFrame.HSTPMenuGroup("My Recent", [], -1)
        r = []
        for program, count in c.most_common(5):
            r.append(BaseAuiFrame.HSTPMenuItem(program, self))
        group.SetSubItems([r])
        return group

    def CreateHTMLWindow(self):
        # Set up the contact tree control with log window
        self.htmlpanel = wx.Panel(self, -1)
        sizer = wx.BoxSizer(wx.VERTICAL)
        btnSizer = wx.FlexGridSizer(1, 4, 10)
        btnSizer.SetFlexibleDirection(wx.HORIZONTAL)

        self.runbtn = btn = wx.Button(self.htmlpanel, -1, "Run Program", style=wx.BU_EXACTFIT)
        self.Bind(wx.EVT_BUTTON, self.OnRunProgram, btn)
        btnSizer.Add(btn, 0, wx.EXPAND | wx.ALL, 2)

        self.addToStartbtn = btn = wx.Button(self.htmlpanel, -1, "Add to Start Menu", style=wx.BU_EXACTFIT)
        self.Bind(wx.EVT_BUTTON, self.OnCreateStartMenuIcon, btn)
        btnSizer.Add(btn, 0, wx.EXPAND | wx.ALL, 2)

        self.addToDesktopbtn = btn = wx.Button(self.htmlpanel, -1, "Add to Desktop", style=wx.BU_EXACTFIT)
        self.Bind(wx.EVT_BUTTON, self.OnCreateDesktopIcon, btn)
        btnSizer.Add(btn, 0, wx.EXPAND | wx.ALL, 2)

        self.debugbtn = btn = wx.Button(self.htmlpanel, -1, "Debug", style=wx.BU_EXACTFIT)
        self.Bind(wx.EVT_BUTTON, self.OnDebugProgram, btn)
        btnSizer.Add(btn, 0, wx.ALL | wx.ALIGN_RIGHT, 2)
        # FIXME @TODO -- this AddGrowableCol is causing a hard crash -- is it the wx version or an interaction with another module?
        # btnSizer.AddGrowableCol(btnSizer.GetItemCount() - 1)

        btnSizer2nd = wx.BoxSizer(wx.HORIZONTAL)
        btn = wx.Button(self.htmlpanel, -1, "<--", style=wx.BU_EXACTFIT)
        self.Bind(wx.EVT_BUTTON, self.OnPrevPageButton, btn)
        btnSizer2nd.Add(btn, 0, wx.EXPAND | wx.ALL, 2)
        self.Bind(wx.EVT_UPDATE_UI, self.OnCheckCanGoBack, btn)

        btn = wx.Button(self.htmlpanel, -1, "-->", style=wx.BU_EXACTFIT)
        self.Bind(wx.EVT_BUTTON, self.OnNextPageButton, btn)
        btnSizer2nd.Add(btn, 0, wx.EXPAND | wx.ALL, 2)
        self.Bind(wx.EVT_UPDATE_UI, self.OnCheckCanGoForward, btn)

        # btn = wx.Button(self.htmlpanel, -1, "Stop", style=wx.BU_EXACTFIT)
        # self.Bind(wx.EVT_BUTTON, self.OnStopButton, btn)
        # btnSizer2nd.Add(btn, 0, wx.EXPAND|wx.ALL, 2)

        try:
            self.htmlview = iewin.IEHtmlWindow(self.htmlpanel)
            self.htmlview.LoadURL = self.htmlview.Navigate
            self.htmlview.AddEventSink(self)
            self.htmlview.SetPage = self.ieSetPage
        except:
            self.htmlview = webview.WebView.New(self.htmlpanel)
            self.htmlview.Bind(webview.EVT_WEBVIEW_NAVIGATING, self.OnWebViewNavigating)
            self.htmlview.Bind(webview.EVT_WEBVIEW_LOADED, self.OnWebViewLoaded)

        self.current_url = "http://wxPython.org"
        self._mgr.AddPane(self.htmlpanel, wx.aui.AuiPaneInfo().
                          Name("Description").Caption("Description").
                          Top().Layer(0).CloseButton(False).MaximizeButton(True).MinSize(wx.Size(400, 300)).FloatingSize((400, 600)))
        self.htmlview.LoadURL(ProgramList["General"].docs)

        sizer.Add(btnSizer, 0, wx.EXPAND)
        sizer.Add(btnSizer2nd, 0, wx.EXPAND)
        sizer.Add(self.htmlview, 1, wx.EXPAND)
        self.htmlpanel.SetSizer(sizer)
        self.ResetButtons(False)

    def ieSetPage(self, html, fake_url=""):
        open(PathToDocs("temp.html"), "w+").write(html)
        self.htmlview.LoadURL(PathToDocs("temp.html"))

    def OnPrevPageButton(self, event):
        self.htmlview.GoBack()

    def OnNextPageButton(self, event):
        self.htmlview.GoForward()

    def OnCheckCanGoBack(self, event):
        event.Enable(self.htmlview.CanGoBack())

    def OnCheckCanGoForward(self, event):
        event.Enable(self.htmlview.CanGoForward())

    def OnStopButton(self, evt):
        self.htmlview.Stop()

    def OnWebViewNavigating(self, evt):
        # this event happens prior to trying to get a resource
        url = evt.GetURL()
        if url.lower().startswith("pydro://"):
            if self.filter.GetValue():
                # clear the tree search or else the highlight can't move to something not displayed
                self.filter.SetValue("")
                self.OnSearch(None)
            program_name = url[8:].replace("%20", " ").replace("/", "")
            children = self.tree.GetItemChildren(self.tree.GetRootItem(), True)
            for c in children:
                if self.tree.GetItemText(c) == program_name:
                    if self.tree.GetItemText(self.tree.GetSelection()) == program_name:
                        self.tree.Unselect()  # remove and reselect to refresh the screen.  If user used back button the tree selection and html window can be mis-matched
                    self.tree.DoSelectItem(c)
                    break
            # self.Launch(program_name)
        if url.lower().startswith("http"):
            self.log.write("Showing " + url + " in external browser\n")
            win32api.ShellExecute(0, None, url, None, '', win32con.SW_SHOW)  # should launch default browser
            evt.Veto()

    def OnWebViewLoaded(self, evt):
        self.current_url = evt.GetURL()
        # self.log.write(evt.GetURL()+"\n")

    def AddMenuItemsToTree(self, items, parentnode, filter_name="", filter_docs=""):
        for i in items:
            if isinstance(i, BaseAuiFrame.HSTPMenuGroup):
                child = self.tree.AppendItem(parentnode, i.GetText())  # , ct_type=2, wnd=self.gauge
                self.AddMenuItemsToTree(i.GetSubItems()[0], child, filter_name=filter_name, filter_docs=filter_docs)
            elif isinstance(i, BaseAuiFrame.HSTPMenuItem):
                mi = i.GetMethodName()
                itemText = i.GetText()
                show = False
                if not filter_name and not filter_docs:
                    show = True
                if not show and filter_name:
                    itxt = itemText.replace("_", " ").lower()
                    if all(x in itxt for x in filter_name.lower().split(" ")):
                        show = True
                if not show and filter_docs:
                    try:
                        docs = open(ProgramList[itemText].docs, "r").read().lower()
                        # remove the table of contents from the search since it lists all the program names
                        try:
                            docs = docs[:docs.find("sphinxsidebar")] + docs[docs.find('id="searchbox"'):]
                        except:
                            pass
                        if all(x in docs for x in filter_docs.lower().split(" ")):
                            show = True
                    except:
                        pass  # no program docs
                # if not show and (not filter_docs or filter_name.lower() in p.replace("_", " ").lower()):
                #    show = True
                if show:
                    child = self.tree.AppendItem(parentnode, i.GetText())  # , ct_type=2, wnd=self.gauge
                    try:
                        self.tree.SetItemImage(child, IconNumbers[i.GetText()], wx.TreeItemIcon_Normal)
                    except:
                        pass
                    part = functools.partial(self.Launch, i.GetText())  # e.g. self.OnSave(self._forms['sample']...)
                    self.__setattr__(mi, part)  # e.g. self.OnSaveSample = self.OnSave(self._forms['sample']...)
                    # print "self."+mi, eval("self."+mi)

    def OnLeftDClick(self, event):
        pt = event.GetPosition()
        item, flags = self.tree.HitTest(pt)
        if item and (flags & CT.TREE_HITTEST_ONITEMLABEL):
            self.Launch(self.tree.GetItemText(item))
        event.Skip()

    def OnItemExpanded(self, event):
        item = event.GetItem()
        wx.LogMessage("OnItemExpanded: %s" % self.tree.GetItemText(item))
        event.Skip()

    # ---------------------------------------------
    def OnItemCollapsed(self, event):
        item = event.GetItem()
        wx.LogMessage("OnItemCollapsed: %s" % self.tree.GetItemText(item))
        event.Skip()

    # ---------------------------------------------
    def OnTreeLeftDown(self, event):
        # reset the overview text if the tree item is clicked on again
        # pt = event.GetPosition();
        # item, flags = self.tree.HitTest(pt)
        # if item == self.tree.GetSelection():
        #     self.log.write("OnLeft: %s\n"%self.tree.GetItemText(item))
        #     #self.SetOverview(self.tree.GetItemText(item)+" Overview", self.curOverview)
        event.Skip()

    def ResetButtons(self, bEnable):
        self.runbtn.Enable(bEnable)
        self.addToStartbtn.Enable(bEnable)
        self.addToDesktopbtn.Enable(bEnable)
        self.debugbtn.Enable(bEnable)

    # ---------------------------------------------
    def OnSelChanged(self, event):
        # if self.dying or not self.loaded or self.skipLoad:
        #    return

        # self.StopDownload()

        item = event.GetItem()
        itemText = self.tree.GetItemText(item)
        try:
            self.tree.Expand(item)
        except:
            pass
        # self.log.write("OnSelection: %s\n"%self.tree.GetItemText(item))
        try:
            self.htmlview.LoadURL(ProgramList[itemText].docs)
            if not (ProgramList[itemText].opts.args or
                    ProgramList[itemText].opts.cmd or
                    ProgramList[itemText].opts.env):
                self.ResetButtons(False)  # no program to run
            else:
                self.ResetButtons(True)

        except KeyError:
            self.ResetButtons(False)
            self.htmlview.LoadURL(ProgramList["General"].docs)
        except:
            if not ProgramList[itemText].docs:
                self.htmlview.SetPage("<b>No Description Found</b><br>This should be showing documentation, but something is missing..", "file://testing.html")
            else:
                se = traceback.format_exc()
                self.htmlview.SetPage("<b>FAILED TO LOAD DOCUMENTATION</b><br>From- %s<br><br> " % ProgramList[itemText].docs + se.replace("\n", "<br><br>"), "file://testing.html")

        # self.LoadDemo(itemText)
        # self.StartDownload()
        event.Skip()


class DemoApp(BaseAuiFrame.SplashScreenApp):

    def ShowMain(self):
        frame = XmlDRFrame(None, -1, "Pydro Explorer v%s" % _PydroVersion)  # wxDefaultPosition, wxSize?
        frame.Show(True)

        return True


def main():
    parser = argparse.ArgumentParser(description="Start the Pydro Explorer app")
    parser.add_argument("-?", "--show_help", action="store_true", help="show this help message and exit")
    parser.add_argument("-d", "--docs", action="store_true", help="Build the program_list_auto.rst file for with links to all included docs")

    args = parser.parse_args()
    if args.show_help:
        parser.print_help()
        sys.exit()

    if args.docs:
        app = wx.App(redirect=False)
        app.MainLoop()
        frame = XmlDRFrame(None, -1, "Pydro Explorer v%s" % _PydroVersion)  # wxDefaultPosition, wxSize?
        frame.MakeRST()
    else:
        app = DemoApp(redirect=False)
        app.MainLoop()


if __name__ == '__main__':
    main()
