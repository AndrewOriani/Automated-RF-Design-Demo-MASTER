'''
pyEPR.ansys
    2014-present

Purpose:
    Handles Ansys interaction and control from version 2014 onward.
    Tested most extensively with V2016 and V2019R3.

@authors:
    Originally contributed by Phil Reinhold.
    Developed further by Zlatko Minev, Zaki Leghtas, and the pyEPR team.
    For the base version of hfss.py, see https://github.com/PhilReinhold/pyHFSS
'''

# Python 2.7 and 3 compatibility
from __future__ import (division, print_function)

from typing import List

import atexit
import os
import re
import signal
import tempfile
import time
import types
import warnings
import glob
import ezdxf
from asteval import Interpreter
from collections.abc import Iterable
from copy import copy
from numbers import Number
from pathlib import Path

import numpy as np
import pandas as pd
from sympy.parsing import sympy_parser
import io

from . import logger

# Handle a  few usually troublesome to import packages, which the use may not have
# installed yet
try:
    import pythoncom
except (ImportError, ModuleNotFoundError):
    pass

try:
    # TODO: Replace `win32com` with Linux compatible package.
    # See Ansys python files in IronPython internal.
    from win32com.client import Dispatch, CDispatch
except (ImportError, ModuleNotFoundError):
    pass

try:
    from pint import UnitRegistry
    ureg = UnitRegistry()
    Q = ureg.Quantity
except(ImportError, ModuleNotFoundError):
    ureg = "Pint module not installed. Please install."


##############################################################################
###

BASIS_ORDER = {"Zero Order": 0,
               "First Order": 1,
               "Second Order": 2,
               "Mixed Order": -1}

# UNITS
# LENGTH_UNIT         --- HFSS UNITS
# #Assumed default input units for ansys hfss
LENGTH_UNIT = 'meter'
# LENGTH_UNIT_ASSUMED --- USER UNITS
# if a user inputs a blank number with no units in `parse_fix`,
# we can assume the following using
LENGTH_UNIT_ASSUMED = 'mm'
#setup a REGEX for finding unit numbers in strings
RE_COMP = re.compile('\d*\.?\d+\s*?(?:[a-zA-Z]+)')
RE_UNIT = re.compile('([a-z]+)')
RE_VAL = re.compile('([0-9]+)')
#setup an asteval interpreter
aeval=Interpreter()


def simplify_arith_expr(expr):
    try:
        out = repr(sympy_parser.parse_expr(str(expr)))
        return out
    except:
        print("Couldn't parse", expr)
        raise

def increment_name(base, existing):
    if not base in existing:
        return base
    n = 1
    def make_name(): return base + str(n)
    while make_name() in existing:
        n += 1
    return make_name()

def extract_value_unit(expr, units):
    """
    :type expr: str
    :type units: str
    :return: float
    """
    try:
        return Q(expr).to(units).magnitude
    except Exception:
        try:
            return float(expr)
        except Exception:
            return expr

def unit_check(string):
    comp=variable_split(string)
    units=[]
    for val in comp:
        units.append(RE_UNIT.findall(val)[0])
    return units

def extract_value_dim(expr):
    """
    type expr: str
    """
    return str(Q(expr).dimensionality)


def parse_entry(entry, convert_to_unit=LENGTH_UNIT):
    '''
    Should take a list of tuple of list... of int, float or str...
    For iterables, returns lists
    '''
    if not isinstance(entry, list) and not isinstance(entry, tuple):
        return extract_value_unit(entry, convert_to_unit)
    else:
        entries = entry
        _entry = []
        for entry in entries:
            _entry.append(parse_entry(entry, convert_to_unit=convert_to_unit))
        return _entry


def fix_units(x, unit_assumed=None):
    '''
    Convert all numbers to string and append the assumed units if needed.
    For an itterable, returns a list
    '''
    unit_assumed = LENGTH_UNIT_ASSUMED if unit_assumed is None else unit_assumed
    if isinstance(x, str):
        # Check if there are already units defined, assume of form 2.46mm  or 2.0 or 4.
        if x[-1].isdigit() or x[-1] == '.':  # number
            return x + unit_assumed
        else:  # units are already appleid
            return x

    elif isinstance(x, Number):
        return fix_units(str(x)+unit_assumed, unit_assumed=unit_assumed)

    elif isinstance(x, Iterable):  # hasattr(x, '__iter__'):
        return [fix_units(y, unit_assumed=unit_assumed) for y in x]
    else:
        return x


def parse_units(x):
    '''
    Convert number, string, and lists/arrays/tuples to numbers scaled
    in HFSS units.

    Converts to                  LENGTH_UNIT = meters  [HFSS UNITS]
    Assumes input units  LENGTH_UNIT_ASSUMED = mm      [USER UNITS]

    [USER UNITS] ----> [HFSS UNITS]
    '''
    return parse_entry(fix_units(x))


def unparse_units(x):
    '''
        Undo effect of parse_unit.

        Converts to     LENGTH_UNIT_ASSUMED = mm     [USER UNITS]
        Assumes input units     LENGTH_UNIT = meters [HFSS UNITS]

        [HFSS UNITS] ----> [USER UNITS]
    '''
    return parse_entry(fix_units(x, unit_assumed=LENGTH_UNIT), LENGTH_UNIT_ASSUMED)


def parse_units_user(x):
    '''
        Convert from user assuemd units to user assumed units
        [USER UNITS] ----> [USER UNITS]
    '''
    return parse_entry(fix_units(x, LENGTH_UNIT_ASSUMED), LENGTH_UNIT_ASSUMED)


def check_path(path):
    check=glob.glob(path)
    if check==[]:
        return False
    else:
        return True

def variable_split(var_str):
    return RE_COMP.findall(var_str)


class VariableString(str):
    def __add__(self, other):
        return var("(%s) + (%s)" % (self, other))

    def __radd__(self, other):
        return var("(%s) + (%s)" % (other, self))

    def __sub__(self, other):
        return var("(%s) - (%s)" % (self, other))

    def __rsub__(self, other):
        return var("(%s) - (%s)" % (other, self))

    def __mul__(self, other):
        return var("(%s) * (%s)" % (self, other))

    def __rmul__(self, other):
        return var("(%s) * (%s)" % (other, self))

    def __div__(self, other):
        return var("(%s) / (%s)" % (self, other))

    def __rdiv__(self, other):
        return var("(%s) / (%s)" % (other, self))

    def __truediv__(self, other):
        return var("(%s) / (%s)" % (self, other))

    def __rtruediv__(self, other):
        return var("(%s) / (%s)" % (other, self))

    def __pow__(self, other):
        return var("(%s) ^ (%s)" % (self, other))

    def __rpow__(self, other):
        return var("(%s) ^ (%s)" % (other, self))

    def __neg__(self):
        return var("-(%s)" % self)

    def __abs__(self):
        return var("abs(%s)" % self)


def var(x):
    if isinstance(x, str):
        return VariableString(x)
    return x


_release_fns = []


def _add_release_fn(fn):
    global _release_fns
    _release_fns.append(fn)
    atexit.register(fn)
    signal.signal(signal.SIGTERM, fn)
    signal.signal(signal.SIGABRT, fn)


def release():
    '''
    Release COM connection to HFSS.
    '''
    global _release_fns
    for fn in _release_fns:
        fn()
    time.sleep(0.1)

    # Note that _GetInterfaceCount is a memeber
    refcount = pythoncom._GetInterfaceCount()  # pylint: disable=no-member

    if refcount > 0:
        print("Warning! %d COM references still alive" % (refcount))
        print("HFSS will likely refuse to shut down")


class COMWrapper(object):
    def __init__(self):
        _add_release_fn(self.release)

    def release(self):
        for k, v in self.__dict__.items():
            if isinstance(v, CDispatch):
                setattr(self, k, None)


class HfssPropertyObject(COMWrapper):
    prop_holder = None
    prop_tab = None
    prop_server = None


def make_str_prop(name, prop_tab=None, prop_server=None):
    return make_prop(name, prop_tab=prop_tab, prop_server=prop_server)


def make_int_prop(name, prop_tab=None, prop_server=None):
    return make_prop(name, prop_tab=prop_tab, prop_server=prop_server, prop_args=["MustBeInt:=", True])


def make_float_prop(name, prop_tab=None, prop_server=None):
    return make_prop(name, prop_tab=prop_tab, prop_server=prop_server, prop_args=["MustBeInt:=", False])


def make_prop(name, prop_tab=None, prop_server=None, prop_args=None):
    def set_prop(self, value, prop_tab=prop_tab, prop_server=prop_server, prop_args=prop_args):
        prop_tab = self.prop_tab if prop_tab is None else prop_tab
        prop_server = self.prop_server if prop_server is None else prop_server
        if isinstance(prop_tab, types.FunctionType):
            prop_tab = prop_tab(self)
        if isinstance(prop_server, types.FunctionType):
            prop_server = prop_server(self)
        if prop_args is None:
            prop_args = []
        self.prop_holder.ChangeProperty(
            ["NAME:AllTabs",
             ["NAME:"+prop_tab,
              ["NAME:PropServers", prop_server],
              ["NAME:ChangedProps",
               ["NAME:"+name, "Value:=", value] + prop_args]]])

    def get_prop(self, prop_tab=prop_tab, prop_server=prop_server):
        prop_tab = self.prop_tab if prop_tab is None else prop_tab
        prop_server = self.prop_server if prop_server is None else prop_server
        if isinstance(prop_tab, types.FunctionType):
            prop_tab = prop_tab(self)
        if isinstance(prop_server, types.FunctionType):
            prop_server = prop_server(self)
        return self.prop_holder.GetPropertyValue(prop_tab, prop_server, name)

    return property(get_prop, set_prop)


def set_property(prop_holder, prop_tab, prop_server, name, value, prop_args=None):
    '''
    More general non obj oriented, functionatl verison
    prop_args = [] by default
    '''
    if not isinstance(prop_server, list):
        prop_server = [prop_server]
    return prop_holder.ChangeProperty(
        ["NAME:AllTabs",
         ["NAME:"+prop_tab,
          ["NAME:PropServers", *prop_server],
          ["NAME:ChangedProps",
           ["NAME:"+name, "Value:=", value] + (prop_args or [])]]])


class HfssApp(COMWrapper):
    def __init__(self, ProgID='AnsoftHfss.HfssScriptInterface'):
        '''
         Connect to IDispatch-based COM object.
             Parameter is the ProgID or CLSID of the COM object.
             This is found in the regkey.

         Version changes for Ansys HFSS for the main object
             v2016 - 'Ansoft.ElectronicsDesktop'
             v2017 and subsequent - 'AnsoftHfss.HfssScriptInterface'

        '''
        super(HfssApp, self).__init__()
        self._app = Dispatch(ProgID)

    def get_app_desktop(self):
        return HfssDesktop(self, self._app.GetAppDesktop())
        # in v2016, there is also getApp - which can be called with HFSS


class HfssDesktop(COMWrapper):
    def __init__(self, app, desktop):
        """
        :type app: HfssApp
        :type desktop: Dispatch
        """
        super(HfssDesktop, self).__init__()
        self.parent = app
        self._desktop = desktop

        # ansys version, needed to check for command changes,
        # since some commands have changed over the years
        self.version = self.get_version()

    def close_all_windows(self):
        self._desktop.CloseAllWindows()

    def project_count(self):
        return self._desktop.Count()

    def get_active_project(self):
        return HfssProject(self, self._desktop.GetActiveProject())

    def get_projects(self):
        return [HfssProject(self, p) for p in self._desktop.GetProjects()]

    def get_project_names(self):
        return self._desktop.GetProjectList()

    def close_project(self, name, save=True):
        if name in list(self.get_project_names()):
            pass
        else:
            warnings.warn('No project of name %s found.'%name)
            return 
        
        if save==True:
            for projs in self.get_projects():
                if projs.name==name:
                    projs.save()
                else:
                    pass
        elif save==False:
            pass
        else:
            raise Exception('ERROR: Save state must be boolean')

        self._desktop.CloseProject(name)   


    def get_messages(self, project_name="", design_name="", level=0):
        """Use:  Collects the messages from a specified project and design.
        Syntax:              GetMessages <ProjectName>, <DesignName>, <SeverityName>
        Return Value:    A simple array of strings.

        Parameters:
        <ProjectName>
            Type:<string>
            Name of the project for which to collect messages.
            An incorrect project name results in no messages (design is ignored)
            An empty project name results in all messages (design is ignored)

        <DesignName>
            Type: <string>
            Name of the design in the named project for which to collect messages
            An incorrect design name results in no messages for the named project
            An empty design name results in all messages for the named project

        <SeverityName>
            Type: <integer>
            Severity is 0-3, and is tied in to info/warning/error/fatal types as follows:
                0 is info and above
                1 is warning and above
                2 is error and fatal
                3 is fatal only (rarely used)
        """
        return self._desktop.GetMessages(project_name, design_name, level)

    def get_version(self):
        return self._desktop.GetVersion()

    def new_project(self):
        return HfssProject(self, self._desktop.NewProject())

    def open_project(self, path):
        ''' returns error if already open '''
        project_name=path.split('\\')[-1].split('.')[0]
        if check_path(path):
            if project_name in self.get_project_names():
                for projs in self.get_projects():
                    if projs.name==project_name:
                        project=projs
                    else:
                        pass
            else:
                project=HfssProject(self, self._desktop.OpenProject(path))
            return project
        else:
            warnings.warn('No project in path of name: %s'%path.split('\\')[-1])
            return None


    def set_active_project(self, name):
        self._desktop.SetActiveProject(name)

    @property
    def project_directory(self):
        return self._desktop.GetProjectDirectory()

    @project_directory.setter
    def project_directory(self, path):
        self._desktop.SetProjectDirectory(path)

    @property
    def library_directory(self):
        return self._desktop.GetLibraryDirectory()

    @library_directory.setter
    def library_directory(self, path):
        self._desktop.SetLibraryDirectory(path)

    @property
    def temp_directory(self):
        return self._desktop.GetTempDirectory()

    @temp_directory.setter
    def temp_directory(self, path):
        self._desktop.SetTempDirectory(path)


class HfssProject(COMWrapper):
    def __init__(self, desktop, project):
        """
        :type desktop: HfssDesktop
        :type project: Dispatch
        """
        super(HfssProject, self).__init__()
        self.parent = desktop
        self._project = project
        #self.name = project.GetName()
        
        #add a materials manager: Andrew Oriani 9/24/2020
        self._material_mgr=self._project.GetDefinitionManager().GetManager("Material")

        self._ansys_version = self.parent.version

    def close(self):
        self._project.Close()

    def make_active(self):
        self.parent.set_active_project(self.name)

    def get_designs(self):
        return [HfssDesign(self, d) for d in self._project.GetDesigns()]
    
    def get_design_names(self):
        names=[]
        for designs in self.get_designs():
            names.append(designs.name)
        return names

    def delete_design(self, name):
        if name in self.get_design_names():
            pass
        else:
            warnings.warn('Unable to find design of name: %s'%name)
            return None
        self._project.DeleteDesign(name)

    def save(self, path=None):
        if path is None:
            self._project.Save()
        else:
            if path.split('.')[-1]!='aedt':
                raise Exception('ERROR: Must be .aedt filetype. Check suffix.')
            else:
                pass
            self._project.SaveAs(path, True)

    def simulate_all(self):
        self._project.SimulateAll()

    def import_dataset(self, path):
        self._project.ImportDataset(path)

    def rename_design(self, design, rename):
        if design in self.get_designs():
            design.rename_design(design.name, rename)
        else:
            raise ValueError('%s design does not exist' % design.name)

    def duplicate_design(self, target, source):
        src_design = self.get_design(source)
        return src_design.duplicate(name=target)

    def get_variable_names(self):
        return [VariableString(s) for s in self._project.GetVariables()]

    def get_variables(self):
        """ Returns the project variables only, which start with $. These are global variables. """
        return {VariableString(s): self.get_variable_value(s) for s in self._project.GetVariables()}

    def get_variable_value(self, name):
        return self._project.GetVariableValue(name)

    def create_variable(self, name, value):
        self._project.ChangeProperty(
            ["NAME:AllTabs",
             ["NAME:ProjectVariableTab",
              ["NAME:PropServers", "ProjectVariables"],
              ["Name:NewProps",
               ["NAME:" + name,
                "PropType:=", "VariableProp",
                "UserDef:=", True,
                "Value:=", value]]]])

    def set_variable(self, name, value):
        if name not in self._project.GetVariables():
            self.create_variable(name, value)
        else:
            self._project.SetVariableValue(name, value)
        return VariableString(name)

    def get_path(self):
        if self._project:
            return self._project.GetPath()
        else:
            raise Exception('''Error: HFSS Project does not have a path.
        Either there is no HFSS project open, or it is not saved.''')

    def new_design(self, name, type):
        name = increment_name(name, [d.GetName()
                                     for d in self._project.GetDesigns()])
        return HfssDesign(self, self._project.InsertDesign("HFSS", name, type, ""))

    def get_design(self, name):
        return HfssDesign(self, self._project.GetDesign(name))

    def get_active_design(self):
        d = self._project.GetActiveDesign()
        if d is None:
            raise EnvironmentError("No Design Active")
        return HfssDesign(self, d)

    def set_active_design(self, name):
        if name in self.get_design_names():
            self._project.SetActiveDesign(name)
        else:
            warnings.warn('Unable to find design: %s'%name)

    def new_dm_design(self, name):
        return self.new_design(name, "DrivenModal")

    def new_em_design(self, name):
        return self.new_design(name, "Eigenmode")

    def new_q3d_design(self, name):
        return HfssDesign(self, self._project.InsertDesign("Q3D Extractor", name, "", ""))

    def get_material_props(self, name):
        props=list(self._material_mgr.GetProperties(name))
        return props
    
    def add_material(self, params, name='UserMaterial'):

        #Added by Andrew Oriani 09/25/2020
        props=self.get_material_props(name)
        if props!=[]:
            name=increment_name(name, name)

        default_props={'permeability': '1.000021',
                        'conductivity': '38000000',
                        'thermal_conductivity': '237.5',
                        'mass_density': '2689',
                        'specific_heat': '951',
                        'youngs_modulus': '69000000000',
                        'poissons_ratio': '0.31',
                        'thermal_expansion_coeffcient': '2.33e-005'}

        param_dict={}
        if type(params)==list:
            for key in list(params)[::2]:
                if key in default_props:
                    param_dict[key]=str(list(params)[params.index(key)+1])
                else:
                    print('%s parameter not valid'%key)
        elif type(params)==dict:
            for key in iter(params.keys()):
                if key in default_props:
                    param_dict[key]=str(params[key])
                else:
                    print('%s parameter not valid'%key)
        else:
            raise Exception('ERROR: Params must be default property list or dict object.')
                    
        params=["NAME:"+name]
        for key in iter(param_dict.keys()):
            params.append(key+":=")
            params.append(param_dict[key])
        
        self._material_mgr.Add(params)
        
    def edit_material(self, params, name):

        #Added by Andrew Oriani 09/26/2020
        props=self.get_material_props(name)
        if props==[]:
            self.add_material(params, name)
            return
        else:
            pass
        props_dict={}
        for key in list(props)[::2]:
            props_dict[key]=str(list(props)[props.index(key)+1])

        default_props={'permeability': '1.000021',
                        'conductivity': '38000000',
                        'thermal_conductivity': '237.5',
                        'mass_density': '2689',
                        'specific_heat': '951',
                        'youngs_modulus': '69000000000',
                        'poissons_ratio': '0.31',
                        'thermal_expansion_coeffcient': '2.33e-005'}

        if type(params)==list:
            for key in list(params)[::2]:
                if key in default_props:
                    props_dict[key]=str(list(params)[params.index(key)+1])
                else:
                    print('%s parameter not valid'%key)
        elif type(params)==dict:
            for key in iter(params.keys()):
                if key in default_props:
                    props_dict[key]=str(params[key])
                else:
                    print('%s parameter not valid'%key)
        else:
            raise Exception('ERROR: Params must be default property list or dict object.')

        params=["NAME:"+name]
        for key in iter(props_dict.keys()):
            params.append(key+":=")
            params.append(props_dict[key])
        
        self._material_mgr.Edit(name, params)

    @property  # v2016
    def name(self):
        return self._project.GetName()


class HfssDesign(COMWrapper):

    def __init__(self, project, design):
        super(HfssDesign, self).__init__()
        self.parent = project
        self._design = design
        self.name = design.GetName()
        self._ansys_version = self.parent._ansys_version

        try:
            # This funciton does not exist if the desing is not HFSS
            self.solution_type = design.GetSolutionType()
        except Exception as e:
            logger.debug(
                f'Exception occured at design.GetSolutionType() {e}. Assuming Q3D design')
            self.solution_type = 'Q3D'

        if design is None:
            return
        self._setup_module = design.GetModule("AnalysisSetup")
        self._solutions = design.GetModule("Solutions")
        self._fields_calc = design.GetModule("FieldsReporter")
        self._output = design.GetModule("OutputVariable")
        self._boundaries = design.GetModule("BoundarySetup")
        self._reporter = design.GetModule("ReportSetup")
        self._modeler = design.SetActiveEditor("3D Modeler")
        self._optimetrics = design.GetModule("Optimetrics")
        self._mesh = design.GetModule("MeshSetup")
        self.modeler = HfssModeler(self)
        self.optimetrics = Optimetrics(self)
        self.setup_args={}


    def make_active(self):
        self.parent.set_active_design(self.name)

    def rename_design(self, name):
        old_name = self._design.GetName()
        self._design.RenameDesignInstance(old_name, name)

    def copy_to_project(self, project):
        project.make_active()
        project._project.CopyDesign(self.name)
        project._project.Paste()
        return project.get_active_design()

    def duplicate(self, name=None):
        dup = self.copy_to_project(self.parent)
        if name is not None:
            dup.rename_design(name)
        return dup

    def get_setup_names(self):
        return self._setup_module.GetSetups()

    def get_setup(self, name=None):
        """
        :rtype: HfssSetup
        """
        setups = self.get_setup_names()
        if not setups:
            raise EnvironmentError(" *** No Setups Present ***")
        if name is None:
            name = setups[0]
        elif name not in setups:
            raise EnvironmentError(
                "Setup {} not found: {}".format(name, setups))

        if self.solution_type == "Eigenmode":
            return HfssEMSetup(self, name)
        elif self.solution_type == "DrivenModal":
            return HfssDMSetup(self, name)
        elif self.solution_type == "Q3D":
            return AnsysQ3DSetup(self, name)

    def create_dm_setup(self, freq_ghz=1, name="Setup", max_delta_s=0.1, max_passes=10,
                        min_passes=1, min_converged=1, pct_refinement=30,
                        basis_order=-1):

        name = increment_name(name, self.get_setup_names())

        setup_args=[
                "NAME:"+name,
                "Frequency:=", str(freq_ghz)+"GHz",
                "MaxDeltaS:=", max_delta_s,
                "MaximumPasses:=", max_passes,
                "MinimumPasses:=", min_passes,
                "MinimumConvergedPasses:=", min_converged,
                "PercentRefinement:=", pct_refinement,
                "IsEnabled:=", True,
                "BasisOrder:=", basis_order
            ]

        self.setup_args[name]=setup_args

        self._setup_module.InsertSetup(
            "HfssDriven", setup_args)
        return HfssDMSetup(self, name)

    def create_em_setup(self, name="Setup", min_freq_ghz=1, n_modes=1, max_delta_f=0.1,
                        max_passes=10, min_passes=1, min_converged=1, pct_refinement=30,
                        basis_order=-1, converge_on_real=True):

        name = increment_name(name, self.get_setup_names())

        setup_args=[
                "NAME:"+name,
                "MinimumFrequency:=", str(min_freq_ghz)+"GHz",
                "NumModes:=", n_modes,
                "MaxDeltaFreq:=", max_delta_f,
                "ConvergeOnRealFreq:=", converge_on_real,
                "MaximumPasses:=", max_passes,
                "MinimumPasses:=", min_passes,
                "MinimumConvergedPasses:=", min_converged,
                "PercentRefinement:=", pct_refinement,
                "IsEnabled:=", True,
                "BasisOrder:=", basis_order,
                "DoLambdaRefine:="	, True,
		        "DoMaterialLambda:="	, True,
		        "SetLambdaTarget:="	, False,
		        "Target:="		, 0.4,
		        "UseMaxTetIncrease:="	, False,
            ]

        self.setup_args[name]=setup_args

        self._setup_module.InsertSetup(
            "HfssEigen", setup_args)
        return HfssEMSetup(self, name)

    def create_q3d_setup(self, name="Setup", adaptive_freq_ghz=1, min_passes=1, max_passes=10,
                        min_converged=1, pct_refinement=30, pct_error=1, soln_order='High', 
                        save_fields=False):

        if self.solution_type!="Q3D":
            raise TypeError('Incorrect solution type: Must be a Q3D Extractor design')

        name=increment_name(name, self.get_setup_names())

        setup_args=["NAME:"+name, 
                  "AdaptiveFreq:=", str(adaptive_freq_ghz)+"GHz", 
                  "EnableDistribProbTypeOption:=", False, 
                  "SaveFields:=", save_fields, 
                  "Enabled:=", True, 
                      ["NAME:Cap", 
                        "MaxPass:=", max_passes, 
                        "MinPass:=", min_passes, 
                        "MinConvPass:=", min_converged, 
                        "PerError:=", pct_error, 
                        "PerRefine:=", pct_refinement, 
                        "AutoIncreaseSolutionOrder:=", False, 
                        "SolutionOrder:=", soln_order], 
                      ]

        self.setup_args[name]=setup_args

        self._setup_module.InsertSetup("Matrix", setup_args)

        return AnsysQ3DSetup(self, name)

    def delete_setup(self, name):
        if name in self.get_setup_names():
            self._setup_module.DeleteSetups(name)

    def delete_full_variation(self, DesignVariationKey="All", del_linked_data=False):
        """
        DeleteFullVariation
        Use:                   Use to selectively make deletions or delete all solution data.
        Command:         HFSS>Results>Clean Up Solutions...
        Syntax:              DeleteFullVariation Array(<parameters>), boolean
        Parameters:      All | <DataSpecifierArray>
                        If, All, all data of existing variations is deleted.
                        Array(<DesignVariationKey>, )
                        <DesignVariationKey>
                            Type: <string>
                            Design variation string.
                        <Boolean>
                        Type: boolean
                        Whether to also delete linked data.
        """
        self._design.DeleteFullVariation("All", False)

    def get_nominal_variation(self):
        """
        Use: Gets the nominal variation string
        Return Value: Returns a string representing the nominal variation
        Returns string such as "Height='0.06mm' Lj='13.5nH'"
        """
        return self._design.GetNominalVariation()

    def create_variable(self, name, value, postprocessing=False):
        if postprocessing == True:
            variableprop = "PostProcessingVariableProp"
        else:
            variableprop = "VariableProp"

        self._design.ChangeProperty(
            ["NAME:AllTabs",
             ["NAME:LocalVariableTab",
              ["NAME:PropServers", "LocalVariables"],
              ["Name:NewProps",
               ["NAME:" + name,
                "PropType:=", variableprop,
                "UserDef:=", True,
                "Value:=", value]]]])

        return VariableString(name)

    def _variation_string_to_variable_list(self, variation_string: str, for_prop_server=True):
        """Example:
            Takes
                "Cj='2fF' Lj='13.5nH'"
            for for_prop_server=True into
                [['NAME:Cj', 'Value:=', '2fF'], ['NAME:Lj', 'Value:=', '13.5nH']]
            or for for_prop_server=False into
                [['Cj', '2fF'], ['Lj', '13.5nH']]
        """
        s = variation_string
        s = s.split(' ')
        s = [s1.strip().strip("''").split("='") for s1 in s]

        if for_prop_server:
            local, project = [], []

            for arr in s:
                to_add = [f'NAME:{arr[0]}', "Value:=",  arr[1]]
                if arr[0][0] == '$':
                    project += [to_add]  # global variable
                else:
                    local += [to_add]  # local variable

            return local, project

        else:
            return s

    def set_variables(self, variation_string: str):
        """
        Set all variables to match a solved variaiton string.

        Args:
            variation_string (str) :  Variaiton string such as
                "Cj='2fF' Lj='13.5nH'"
        """
        assert isinstance(variation_string, str)

        content = ["NAME:ChangedProps"]
        local, project = self._variation_string_to_variable_list(
            variation_string)
        #print('\nlocal=', local, '\nproject=', project)

        if len(project) > 0:
            self._design.ChangeProperty(
                ["NAME:AllTabs",
                    ["NAME:ProjectVariableTab",
                        ["NAME:PropServers",
                         "ProjectVariables"
                         ],
                        content + project
                     ]
                 ])

        if len(local) > 0:
            self._design.ChangeProperty(
                ["NAME:AllTabs",
                    ["NAME:LocalVariableTab",
                        ["NAME:PropServers",
                         "LocalVariables"
                         ],
                        content + local
                     ]
                 ])

    def set_variable(self, name: str, value: str, postprocessing=False):
        """Warning: THis is case sensitive,

        Arguments:
            name {str} -- Name of variable to set, such as 'Lj_1'.
                          This is not the same as as 'LJ_1'.
                          You must use the same casing.
            value {str} -- Value, such as '10nH'

        Keyword Arguments:
            postprocessing {bool} -- Postprocessingh variable only or not.
                          (default: {False})

        Returns:
            VariableString
        """
        # TODO: check if variable does not exist and quit if it doesn't?
        if name not in self.get_variable_names():
            self.create_variable(name, value, postprocessing=postprocessing)
        else:
            self._design.SetVariableValue(name, value)

        return VariableString(name)

    def get_variable_value(self, name):
        """ Can only access the design variables, i.e., the local ones
            Cannot access the project (global) variables, which start with $. """
        return self._design.GetVariableValue(name)

    def _variable_eval(self, vars):
        model = self.modeler
        var_names = self.get_variable_names()
        if [f for f in vars.split(self._check_for_variable(vars)[0]) if f!='']==[]:
            var_val=self.get_variable_value(vars)
        else:
            var_val=vars
        
        used_vars=self._check_for_variable(var_val)
        if used_vars!=[]:
            for variable in used_vars:
                variable_val=self.get_variable_value(variable)
                sub_vars=[names for names in var_names if names in variable_val]
                if sub_vars!=[]:
                    aeval.symtable[variable]=self._variable_eval(variable)
                else:
                    aeval.symtable[variable]=parse_units(self.get_variable_value(variable))
        else:
            pass
        
        string_search=variable_split(var_val)
        for val in string_search:
            var_val=var_val.replace(val, str(parse_units(val)))
        model_units=model.get_units()
        return aeval.eval(var_val)

    def _check_for_variable(self, string):
        return [names for names in self.get_variable_names() if names in string]

    def conv_variable_value(self, vars):
        '''
        Flattens variables to from HFSS to 
        '''
        var_names=self.get_variable_names()
        if type(vars)==list:
            if [el for el in vars if type(el)==str or type(el)==VariableString]==[]:
                return parse_units(vars)
            else:
                out_var=[]
                for val in vars:
                    if type(val)!=float and type(val)!=int:
                        if self._check_for_variable(val)!=[]: 
                            out_var.append(self._variable_eval(val))
                        else:
                            out_var.append(parse_units(val))
                    else:
                        out_var.append(parse_units(val))
                return out_var 
        else:
            if type(vars)!=float and type(vars)!=int:
                if self._check_for_variable(vars)!=[]:
                    return self._variable_eval(vars)
                else:
                    return parse_units(vars)
            else:
                return parse_units(vars)

    def get_variable_names(self):
        """ Returns the local design variables.
            Does not return the project (global) variables, which start with $. """
        return [VariableString(s) for s in
                self._design.GetVariables()+self._design.GetPostProcessingVariables()]

    def get_variables(self):
        """ Returns dictionary of local design variables and their values.
            Does not return the project (global) variables and their values,
            whose names start with $. """
        local_variables = self._design.GetVariables(
        )+self._design.GetPostProcessingVariables()
        return {lv: self.get_variable_value(lv) for lv in local_variables}

    def copy_design_variables(self, source_design):
        ''' does not check that variables are all present '''

        # don't care about values
        source_variables = source_design.get_variables()

        for name, value in source_variables.items():
            self.set_variable(name, value)

    def get_excitations(self):
        self._boundaries.GetExcitations()

    def _evaluate_variable_expression(self, expr, units):
        """
        :type expr: str
        :type units: str
        :return: float
        """
        try:
            sexp = sympy_parser.parse_expr(expr)
        except SyntaxError:
            return Q(expr).to(units).magnitude

        sub_exprs = {fs: self.get_variable_value(fs.name)
                     for fs in sexp.free_symbols}

        return float(sexp.subs({fs: self._evaluate_variable_expression(e, units)
                                for fs, e in sub_exprs.items()}))

    def eval_expr(self, expr, units="mm"):
        return str(self._evaluate_variable_expression(expr, units)) + units

    def Clear_Field_Calc_Stack(self):
        self._fields_calc.CalcStack("Clear")


class HfssSetup(HfssPropertyObject):
    prop_tab = "HfssTab"
    passes = make_int_prop("Passes")  # see EditSetup
    n_modes = make_int_prop("Modes")
    pct_refinement = make_float_prop("Percent Refinement")
    delta_f = make_float_prop("Delta F")

    min_freq = make_float_prop("Min Freq")
    basis_order = make_str_prop("Basis Order")

    def __init__(self, design, setup):
        """
        :type design: HfssDesign
        :type setup: Dispatch

        :COM Scripting Help: "Analysis Setup Module Script Commands"

        Get properties:
            setup.parent._design.GetProperties("HfssTab",'AnalysisSetup:Setup1')
        """
        super(HfssSetup, self).__init__()
        self.parent = design
        self.prop_holder = design._design
        self._setup_module = design._setup_module
        self._reporter = design._reporter
        self._solutions = design._solutions
        self.name = setup
        self.solution_name = setup + " : LastAdaptive"
        #self.solution_name_pass = setup + " : AdaptivePass"
        self.prop_server = "AnalysisSetup:" + setup
        self.expression_cache_items = ['NAME:ExpressionCache']
        self._ansys_version = self.parent._ansys_version

    def analyze(self, name=None):
        '''
        Use:             Solves a single solution setup and all of its frequency sweeps.
        Command:         Right-click a solution setup in the project tree, and then click Analyze
                         on the shortcut menu.
        Syntax:          Analyze(<SetupName>)
        Parameters:      <setupName>
        Return Value:    None
        -----------------------------------------------------

        Will block the until the analysis is completly done.
        Will raise a com_error if analysis is aborted in HFSS.
        '''
        if name is None:
            name = self.name
        logger.info(f'Analyzing setup {name}')
        return self.parent._design.Analyze(name)

    def solve(self, name=None):
        '''
        Use:             Performs a blocking simulation.
                         The next script command will not be executed
                         until the simulation is complete.

        Command:         HFSS>Analyze
        Syntax:          Solve <SetupNameArray>
        Return Value:   Type: <int>
                        -1: simulation error
                        0: normal completion
        Parameters:      <SetupNameArray>: Array(<SetupName>, <SetupName>, ...)
           <SetupName>
        Type: <string>
        Name of the solution setup to solve.
        Example:
            return_status = oDesign.Solve Array("Setup1", "Setup2")
        -----------------------------------------------------

        HFSS abort: still returns 0 , since termination by user.

        '''
        if name is None:
            name = self.name
        return self.parent._design.Solve(name)

    def insert_sweep(self, start_ghz, stop_ghz, count=None, step_ghz=None,
                     name="Sweep", type="Fast", save_fields=False):

        if not type in ['Fast', 'Interpolating', 'Discrete']:
            logger.error(
                "insert_sweep: Error type was not in  ['Fast', 'Interpolating', 'Discrete']")

        name = increment_name(name, self.get_sweep_names())
        params = [
            "NAME:"+name,
            "IsEnabled:=", True,
            "Type:=", type,
            "SaveFields:=", save_fields,
            "SaveRadFields:=", False,
            # "GenerateFieldsForAllFreqs:="
            "ExtrapToDC:=", False,
        ]

        # not sure hwen extacyl this changed between 2016 and 2019
        if self._ansys_version >= '2019':
            if count:
                params.extend([
                    "RangeType:=",  'LinearCount',
                    "RangeStart:=", f"{start_ghz:f}GHz",
                    "RangeEnd:=",   f"{stop_ghz:f}GHz",
                    "RangeCount:=", count])
            if step_ghz:
                params.extend([
                    "RangeType:=",  'LinearStep',
                    "RangeStart:=", f"{start_ghz:f}GHz",
                    "RangeEnd:=",   f"{stop_ghz:f}GHz",
                    "RangeStep:=", step_ghz])

            if (count and step_ghz) or ((not count) and (not step_ghz)):
                logger.error('ERROR: you should provide either step_ghz or count \
                    when inserting an HFSS driven model freq sweep. \
                    YOu either provided both or neither! See insert_sweep.')
        else:
            params.extend([
                "StartValue:=", "%fGHz" % start_ghz,
                "StopValue:=", "%fGHz" % stop_ghz])
            if step_ghz is not None:
                params.extend([
                    "SetupType:=", "LinearSetup",
                    "StepSize:=", "%fGHz" % step_ghz])
            else:
                params.extend([
                    "SetupType:=", "LinearCount",
                    "Count:=", count])

        self._setup_module.InsertFrequencySweep(self.name, params)

        return HfssFrequencySweep(self, name)

    def delete_sweep(self, name):
        self._setup_module.DeleteSweep(self.name, name)

    def add_fields_convergence_expr(self, expr, pct_delta, context_line, phase=0, num_pts=101):
        """note: because of hfss idiocy, you must call "commit_convergence_exprs"
            after adding all exprs"""
        assert isinstance(expr, NamedCalcObject)
        assert isinstance(context_line, str)
        if context_line in self.parent.modeler.get_objects_in_group('Lines'):
            pass
        else:
            raise Exception('ERROR: Context line is not valid line object')
        self.expression_cache_items.append(
            ["NAME:CacheItem",
                "Title:=", expr.name+"_conv",
                "Expression:=", expr.name,
                "Intrinsics:=", "Phase='{}deg'".format(phase),
                "IsConvergence:=", True,
                "UseRelativeConvergence:=", 1,
                "MaxConvergenceDelta:=", pct_delta,
                "MaxConvergeValue:=", "0.05",
                "ReportType:=", "Fields",
                [   "NAME:ExpressionContext",
                    "Context:="		, context_line,
				    "PointCount:="		, num_pts]])

    def commit_convergence_exprs(self):
        """note: this will eliminate any convergence expressions not added
            through this interface"""
        args=self.parent.setup_args[self.name]
        args.append(self.expression_cache_items)
        self._setup_module.EditSetup(self.name, args)

    def get_sweep_names(self):
        return self._setup_module.GetSweeps(self.name)

    def get_sweep(self, name=None):
        sweeps = self.get_sweep_names()
        if not sweeps:
            raise EnvironmentError("No Sweeps Present")
        if name is None:
            name = sweeps[0]
        elif name not in sweeps:
            raise EnvironmentError(
                "Sweep {} not found in {}".format(name, sweeps))
        return HfssFrequencySweep(self, name)

    # def add_fields_convergence_expr(self, expr, pct_delta, phase=0):
    #     """note: because of hfss idiocy, you must call "commit_convergence_exprs"
    #     after adding all exprs"""
    #     assert isinstance(expr, NamedCalcObject)
    #     self.expression_cache_items.append(
    #         ["NAME:CacheItem",
    #          "Title:=", expr.name+"_conv",
    #          "Expression:=", expr.name,
    #          "Intrinsics:=", "Phase='{}deg'".format(phase),
    #          "IsConvergence:=", True,
    #          "UseRelativeConvergence:=", 1,
    #          "MaxConvergenceDelta:=", pct_delta,
    #          "MaxConvergeValue:=", "0.05",
    #          "ReportType:=", "Fields",
    #          ["NAME:ExpressionContext"]])

    # def commit_convergence_exprs(self):
    #     """note: this will eliminate any convergence expressions not added through this interface"""
    #     args = [
    #         "NAME:"+self.name,
    #         ["NAME:ExpressionCache", self.expression_cache_items]
    #     ]
    #     self._setup_module.EditSetup(self.name, args)

    def get_convergence(self, variation="", pre_fn_args=[], overwrite=True):
        '''
        Returns converge as a dataframe
            Variation should be in the form
            variation = "scale_factor='1.2001'" ...
        '''
        # TODO: (Daniel) I think this data should be store in a more comfortable datatype (dictionary maybe?)
        # Write file
        temp = tempfile.NamedTemporaryFile()
        temp.close()
        temp = temp.name + '.conv'
        self.parent._design.ExportConvergence(
            self.name, variation, *pre_fn_args, temp, overwrite)

        # Read File
        temp = Path(temp)
        if not temp.is_file():
            logger.error(f'''ERROR!  Error in trying to read temporary convergence file.
                        `get_convergence` did not seem to have the file written {str(temp)}.
                        Perhaps there was no convergence?  Check to see if there is a CONV available for this current variation. If the nominal design is not solved, it will not have a CONV., but will show up as a variation
                        Check for error messages in HFSS.
                        Retuning None''')
            return None, ''
        text = temp.read_text()

        # Parse file
        text2 = text.split(r'==================')
        if len(text) >= 3:
            df = pd.read_csv(io.StringIO(
                text2[3].strip()), sep='|', skipinitialspace=True, index_col=0).drop('Unnamed: 3', 1)
        else:
            logger.error(f'ERROR IN reading in {temp}:\n{text}')
            df = None

        return df, text

    def get_mesh_stats(self, variation=""):
        '''  variation should be in the form
             variation = "scale_factor='1.2001'" ...
        '''
        temp = tempfile.NamedTemporaryFile()
        temp.close()
        # print(temp.name0
        # seems broken in 2016 because of extra text added to the top of the file
        self.parent._design.ExportMeshStats(
            self.name, variation, temp.name + '.mesh', True)
        try:
            df = pd.read_csv(temp.name+'.mesh', delimiter='|', skipinitialspace=True,
                             skiprows=7, skipfooter=1, skip_blank_lines=True, engine='python')
            df = df.drop('Unnamed: 9', 1)
        except Exception as e:
            print("ERROR in MESH reading operation.")
            print(e)
            print('ERROR!  Error in trying to read temporary MESH file ' + temp.name +
                  '\n. Check to see if there is a mesh available for this current variation.\
                   If the nominal design is not solved, it will not have a mesh., \
                   but will show up as a variation.')
            df = None
        return df

    def get_profile(self, variation=""):
        fn = tempfile.mktemp()
        self.parent._design.ExportProfile(self.name, variation, fn, False)
        df = pd.read_csv(fn, delimiter='\t', skipinitialspace=True, skiprows=6,
                         skipfooter=1, skip_blank_lines=True, engine='python')
        # just borken down by new lines
        return df

    def get_fields(self):
        return HfssFieldsCalc(self)


class HfssDMSetup(HfssSetup):
    """
    Driven modal setup
    """
    solution_freq = make_float_prop("Solution Freq")
    delta_s = make_float_prop("Delta S")
    solver_type = make_str_prop("Solver Type")

    def setup_link(self, linked_setup):
        '''
            type: linked_setup <HfssSetup>
        '''
        args = ["NAME:" + self.name,
                ["NAME:MeshLink",
                 "Project:=", "This Project*",
                 "Design:=", linked_setup.parent.name,
                 "Soln:=", linked_setup.solution_name,
                 self._map_variables_by_name(),
                 "ForceSourceToSolve:=", True,
                 "PathRelativeTo:=", "TargetProject",
                 ],
                ]
        self._setup_module.EditSetup(self.name, args)

    def _map_variables_by_name(self):
        ''' does not check that variables are all present '''
        # don't care about values
        project_variables = self.parent.parent.get_variable_names()
        design_variables = self.parent.get_variable_names()

        # build array
        args = ["NAME:Params", ]
        for name in project_variables:
            args.extend([str(name)+":=", str(name)])
        for name in design_variables:
            args.extend([str(name)+":=", str(name)])
        return args

    def get_solutions(self):
        return HfssDMDesignSolutions(self, self.parent._solutions)


class HfssEMSetup(HfssSetup):
    """
    Eigenmode setup
    """
    min_freq = make_float_prop("Min Freq")
    n_modes = make_int_prop("Modes")
    delta_f = make_float_prop("Delta F")

    def get_solutions(self):
        return HfssEMDesignSolutions(self, self.parent._solutions)


class AnsysQ3DSetup(HfssSetup):
    """
    Q3D setup
    """
    prop_tab = "CG"
    max_pass = make_int_prop("Max. Number of Passes")
    max_pass = make_int_prop("Min. Number of Passes")
    pct_error = make_int_prop("Percent Error")
    frequency = make_str_prop("Adaptive Freq", 'General')  # e.g., '5GHz'
    n_modes = 0  # for compatability with eigenmode

    def get_frequency_Hz(self):
        return int(ureg(self.frequency).to('Hz').magnitude)

    def get_solutions(self):
        return HfssQ3DDesignSolutions(self, self.parent._solutions)

    def get_convergence(self, variation=""):
        '''
        Returns df
                    # Triangle   Delta %
            Pass
            1            164       NaN
        '''
        return super().get_convergence(variation, pre_fn_args=['CG'])

    def get_matrix(self, variation='', pass_number=0, frequency=None,
                   MatrixType='Maxwell',
                   solution_kind='LastAdaptive',  # AdpativePass
                   ACPlusDCResistance=False,
                   soln_type="C"):
        '''
        Arguments:
        -----------
            variation : an empty string returns nominal variation.
                        Otherwise need the list
            frequency : in Hz
            soln_type = "C", "AC RL" and "DC RL"
            solution_kind = 'LastAdaptive' # AdaptivePass
        Internals:
        -----------
            Uses self.solution_name  = Setup1 : LastAdaptive

        Returns:
        ---------------------
            df_cmat, user_units, (df_cond, units_cond), design_variation
        '''
        if frequency is None:
            frequency = self.get_frequency_Hz()

        temp = tempfile.NamedTemporaryFile()
        temp.close()
        path = temp.name+'.txt'
        # <FileName>, <SolnType>, <DesignVariationKey>, <Solution>, <Matrix>, <ResUnit>,
        # <IndUnit>, <CapUnit>, <CondUnit>, <Frequency>, <MatrixType>, <PassNumber>,
        # <ACPlusDCResistance>
        self.parent._design.ExportMatrixData(path, soln_type, variation,
                                             f'{self.name}:{solution_kind}',
                                             "Original", "ohm", "nH", "fF", "mSie",
                                             frequency, MatrixType,
                                             pass_number, ACPlusDCResistance)

        df_cmat, user_units, (df_cond, units_cond), design_variation = \
            self.load_q3d_matrix(path)
        return df_cmat, user_units, (df_cond, units_cond), design_variation

    def get_matrix_dict(self, variation='', pass_number=0, frequency=None,
                        MatrixType='Maxwell',
                        solution_kind='LastAdaptive',  # AdpativePass
                        ACPlusDCResistance=False,
                        soln_type="C"):
        
        if frequency is None:
            frequency = self.get_frequency_Hz()

        temp = tempfile.NamedTemporaryFile()
        temp.close()
        path = temp.name+'.txt'
        # <FileName>, <SolnType>, <DesignVariationKey>, <Solution>, <Matrix>, <ResUnit>,
        # <IndUnit>, <CapUnit>, <CondUnit>, <Frequency>, <MatrixType>, <PassNumber>,
        # <ACPlusDCResistance>
        self.parent._design.ExportMatrixData(path, soln_type, variation,
                                             f'{self.name}:{solution_kind}',
                                             "Original", "ohm", "nH", "fF", "mSie",
                                             frequency, MatrixType,
                                             pass_number, ACPlusDCResistance)
        text=Path(path).read_text()
        
        s1 = text.split('Capacitance Matrix')
        assert len(s1) == 2, "Could not split text to `Capacitance Matrix`"

        s2 = s1[1].split('Conductance Matrix')

        
        cap_data=s2[0].strip().split('\n')
        cond_data=s2[1].strip().split('\n')
        row_num=len(cap_data)
        col_num=len(cap_data[0].split('\t'))
        cap_array=np.zeros((row_num-1, col_num-1))
        cond_array=np.zeros((row_num-1, col_num-1))
        
        cond_unit=re.findall(r'G Units:(.*?)\n', text)[0]
        cap_unit=re.findall(r'C Units:(.*?)\n', text)[0].split(',')[0]
        
        net_names=cap_data[0].split('\t')[0:-1]
        
        name_dict={}
        for I, names in enumerate(net_names):
            name_dict[names]=I
        
        for I, rows in enumerate(cap_data[1::]):
            cols=rows.split('\t')
            for J, col in enumerate(cols[1::]):
                cap_array[I, J]=col
                
        for I, rows in enumerate(cond_data[1::]):
            cols=rows.split('\t')
            for J, col in enumerate(cols[1::]):
                cond_array[I, J]=col
                
        data_dict={'capacitance': cap_array, 'conductivity': cond_array}
        unit_dict={'capacitance': cap_unit, 'conductivity': cond_unit}
        
        out_dict={'matrix': data_dict, 'units': unit_dict, 'nets': name_dict, 'frequency': frequency}
        
        return out_dict


    @staticmethod
    def _readin_Q3D_matrix(path):
        """
        Read in the txt file created from q3d export
        and output the capacitance matrix

        When exporting pick "save as type: data table"

        See Zlatko

        RETURNS: Dataframe

        Example file:
        ```
        DesignVariation:$BBoxL='650um' $boxH='750um' $boxL='2mm' $QubitGap='30um' \
                        $QubitH='90um' \$QubitL='450um' Lj_1='13nH'
        Setup1:LastAdaptive
        Problem Type:C
        C Units:farad, G Units:mSie
        Reduce Matrix:Original
        Frequency: 5.5E+09 Hz

        Capacitance Matrix
            ground_plane	Q1_bus_Q0_connector_pad	Q1_bus_Q2_connector_pad	Q1_pad_bot	Q1_pad_top1	Q1_readout_connector_pad
        ground_plane	2.8829E-13	-3.254E-14	-3.1978E-14	-4.0063E-14	-4.3842E-14	-3.0053E-14
        Q1_bus_Q0_connector_pad	-3.254E-14	4.7257E-14	-2.2765E-16	-1.269E-14	-1.3351E-15	-1.451E-16
        Q1_bus_Q2_connector_pad	-3.1978E-14	-2.2765E-16	4.5327E-14	-1.218E-15	-1.1552E-14	-5.0414E-17
        Q1_pad_bot	-4.0063E-14	-1.269E-14	-1.218E-15	9.5831E-14	-3.2415E-14	-8.3665E-15
        Q1_pad_top1	-4.3842E-14	-1.3351E-15	-1.1552E-14	-3.2415E-14	9.132E-14	-1.0199E-15
        Q1_readout_connector_pad	-3.0053E-14	-1.451E-16	-5.0414E-17	-8.3665E-15	-1.0199E-15	3.9884E-14

        Conductance Matrix
            ground_plane	Q1_bus_Q0_connector_pad	Q1_bus_Q2_connector_pad	Q1_pad_bot	Q1_pad_top1	Q1_readout_connector_pad
        ground_plane	0	0	0	0	0	0
        Q1_bus_Q0_connector_pad	0	0	0	0	0	0
        Q1_bus_Q2_connector_pad	0	0	0	0	0	0
        Q1_pad_bot	0	0	0	0	0	0
        Q1_pad_top1	0	0	0	0	0	0
        Q1_readout_connector_pad	0	0	0	0	0	0
        ```
        """

        text = Path(path).read_text()


        s1 = text.split('Capacitance Matrix')
        assert len(s1) == 2, "Could not split text to `Capacitance Matrix`"

        s2 = s1[1].split('Conductance Matrix')

        df_cmat = pd.read_csv(io.StringIO(
            s2[0].strip()), delim_whitespace=True, skipinitialspace=True, index_col=0)
        units = re.findall(r'C Units:(.*?),', text)[0]

        if len(s2) > 1:
            df_cond = pd.read_csv(io.StringIO(
                s2[1].strip()), delim_whitespace=True, skipinitialspace=True, index_col=0)
            units_cond = re.findall(r'G Units:(.*?)\n', text)[0]
        else:
            df_cond = None

        var = re.findall(r'DesignVariation:(.*?)\n', text) # this changed circe v2020
        if len(var) <1: # didnt find
            var = re.findall(r'Design Variation:(.*?)\n', text)
            if len(var) <1: # didnt find
                logger.error(f'Failed to parse Q3D matrix Design Variation:\nFile:{path}\nText:{text}')

                var = ['']
        design_variation = var[0]

        return df_cmat, units, design_variation, df_cond, units_cond

    @staticmethod
    def load_q3d_matrix(path, user_units='fF'):
        """Load Q3D capcitance file exported as Maxwell matrix.
        Exports also conductance conductance.
        Units are read in automatically and converted to user units.

        Arguments:
            path {[str or Path]} -- [path to file text with matrix]

        Returns:
            df_cmat, user_units, (df_cond, units_cond), design_variation

            dataframes: df_cmat, df_cond
        """
        df_cmat, Cunits, design_variation, df_cond, units_cond = AnsysQ3DSetup._readin_Q3D_matrix(
            path)

        # Unit convert
        q = ureg.parse_expression(Cunits).to(user_units)
        df_cmat = df_cmat * q.magnitude  # scale to user units

        #print("Imported capacitance matrix with UNITS: [%s] now converted to USER UNITS:[%s] from file:\n\t%s"%(Cunits, user_units, path))

        return df_cmat, user_units, (df_cond, units_cond), design_variation


class HfssDesignSolutions(COMWrapper):
    def __init__(self, setup, solutions):
        '''
        :type setup: HfssSetup
        '''
        super(HfssDesignSolutions, self).__init__()
        self.parent = setup
        self._solutions = solutions
        self._ansys_version = self.parent._ansys_version

    def get_valid_solution_list(self):
        '''
         Gets all available solution names that exist in a design.
         Return example:
            ('Setup1 : AdaptivePass', 'Setup1 : LastAdaptive')
        '''
        return self._solutions.GetValidISolutionList()

    def list_variations(self, setup_name: str = None):
        """
        Get a list of solved variations.

        Args:
            setup_name(str) : Example name ("Setup1 : LastAdaptive") Defaults to None.

        Returns:
             An array of strings corresponding to solved variations.

             .. code-block:: python

                ("Cj='2fF' Lj='12nH'",
                "Cj='2fF' Lj='12.5nH'",
                "Cj='2fF' Lj='13nH'",
                "Cj='2fF' Lj='13.5nH'",
                "Cj='2fF' Lj='14nH'")
        """
        if setup_name is None:
            setup_name = str(self.parent.solution_name)
        return self._solutions.ListVariations(setup_name)


class HfssEMDesignSolutions(HfssDesignSolutions):

    def eigenmodes(self, lv=""):
        '''
        Returns the eigenmode data of freq and kappa/2p
        '''
        fn = tempfile.mktemp()
        #print(self.parent.solution_name, lv, fn)
        self._solutions.ExportEigenmodes(self.parent.solution_name, lv, fn)
        data = np.genfromtxt(fn, dtype='str')
        # Update to Py 3:
        # np.loadtxt and np.genfromtxt operate in byte mode, which is the default string type in Python 2.
        # But Python 3 uses unicode, and marks bytestrings with this b.
        # getting around the very annoying fact that
        if np.size(np.shape(data)) == 1:
            # in Python a 1D array does not have shape (N,1)
            data = np.array([data])
        else:                                  # but rather (N,) ....
            pass
        if np.size(data[0, :]) == 6:  # checking if values for Q were saved
            # eigvalue=(omega-i*kappa/2)/2pi
            kappa_over_2pis = [2*float(ii) for ii in data[:, 3]]
            # so kappa/2pi = 2*Im(eigvalue)
        else:
            kappa_over_2pis = None

        # print(data[:,1])
        freqs = [float(ii) for ii in data[:, 1]]
        return freqs, kappa_over_2pis

    """
    Export eigenmodes vs pass number
    Did not figre out how to set pass number in a hurry.


    import tempfile
    self = epr_hfss.solutions

    '''
    HFSS: Exports a tab delimited table of Eigenmodes in HFSS. Not in HFSS-IE.
    <setupName> <solutionName> <DesignVariationKey>
    <filename>
    Return Value:    None

    Parameters:
        <SolutionName>
            Type: <string>
            Name of the solutions within the solution setup.
        <DesignVariationKey>
            Type: <string>
            Design variation string.
    '''
    setup = self.parent
    fn = tempfile.mktemp()
    variation_list=''
    soln_name = f'{setup.name} : AdaptivePas'
    available_solns = self._solutions.GetValidISolutionList()
    if not(soln_name in available_solns):
        logger.error(f'ERROR Tried to export freq vs pass number, but solution  `{soln_name}` was not in avaialbe `{available_solns}`. Returning []')
        #return []
    self._solutions.ExportEigenmodes(soln_name, ['Pass:=5'], fn) # ['Pass:=5'] fails  can do with ''
    """

    def set_mode(self, n, phase=0, FieldType='EigenStoredEnergy'):
        '''
        Indicates which source excitations should be used for fields post processing.
        HFSS>Fields>Edit Sources

        Mode count starts at 1

        Amplitude is set to 1

        No error is thorwn if a number exceeding number of modes is set

            FieldType -- EigenStoredEnergy or EigenPeakElecticField
        '''
        n_modes = int(self.parent.n_modes)

        if n < 1:
            err = f'ERROR: You tried to set a mode < 1. {n}/{n_modes}'
            logger.error(err)
            raise Exception(err)

        if n > n_modes:
            err = f'ERROR: You tried to set a mode > number of modes {n}/{n_modes}'
            logger.error(err)
            raise Exception(err)

        if self._ansys_version >= '2019':
            # THIS WORKS FOR v2019R2
            self._solutions.EditSources(
                [
                    [
                        "FieldType:=", "EigenPeakElectricField"
                    ],
                    [
                        "Name:=", "Modes",
                        "Magnitudes:=", ["1" if i + 1 ==
                                         n else "0" for i in range(n_modes)],
                        "Phases:=", [str(phase) if i + 1 ==
                                     n else "0" for i in range(n_modes)]
                    ]
                ])
        else:
            # The syntax has changed for AEDT 18.2.
            # see https://ansyshelp.ansys.com/account/secured?returnurl=/Views/Secured/Electronics/v195//Subsystems/HFSS/Subsystems/HFSS%20Scripting/HFSS%20Scripting.htm

            self._solutions.EditSources(
                "EigenStoredEnergy",
                ["NAME:SourceNames", "EigenMode"],
                ["NAME:Modes", n_modes],
                ["NAME:Magnitudes"] + [1 if i + 1 ==
                                       n else 0 for i in range(n_modes)],
                ["NAME:Phases"] + [phase if i + 1 ==
                                   n else 0 for i in range(n_modes)],
                ["NAME:Terminated"],
                ["NAME:Impedances"]
            )

    def has_fields(self, variation_string=None):
        '''
        Determine if fields exist for a particular solution.

        variation_string : str | None
            This must the string that describes the variaiton in hFSS, not 0 or 1, but
            the string of variables, such as
                "Cj='2fF' Lj='12.75nH'"
            If None, gets the nominal variation
        '''
        if variation_string is None:
            variation_string = self.parent.parent.get_nominal_variation()

        return bool(self._solutions.HasFields(self.parent.solution_name, variation_string))

    def create_report(self, plot_name, xcomp, ycomp, params, pass_name='LastAdaptive'):
        '''
        pass_name: AdaptivePass, LastAdaptive

        Example
        ------------------------------------------------------
        Exammple plot for a single vareiation all pass converge of mode freq
        .. code-block python
            ycomp = [f"re(Mode({i}))" for i in range(1,1+epr_hfss.n_modes)]
            params = ["Pass:=", ["All"]]+variation
            setup.create_report("Freq. vs. pass", "Pass", ycomp, params, pass_name='AdaptivePass')
        '''
        assert isinstance(ycomp, list)
        assert isinstance(params, list)

        setup = self.parent
        reporter = setup._reporter
        return reporter.CreateReport(plot_name, "Eigenmode Parameters", "Rectangular Plot",
                                     f"{setup.name} : {pass_name}", [], params,
                                     ["X Component:=", xcomp,
                                      "Y Component:=", ycomp], [])


class HfssDMDesignSolutions(HfssDesignSolutions):

    pass


class HfssQ3DDesignSolutions(HfssDesignSolutions):
    pass


class HfssFrequencySweep(COMWrapper):
    prop_tab = "HfssTab"
    start_freq = make_float_prop("Start")
    stop_freq = make_float_prop("Stop")
    step_size = make_float_prop("Step Size")
    count = make_float_prop("Count")
    sweep_type = make_str_prop("Type")

    def __init__(self, setup, name):
        """
        :type setup: HfssSetup
        :type name: str
        """
        super(HfssFrequencySweep, self).__init__()
        self.parent = setup
        self.name = name
        self.solution_name = self.parent.name + " : " + name
        self.prop_holder = self.parent.prop_holder
        self.prop_server = self.parent.prop_server + ":" + name
        self._ansys_version = self.parent._ansys_version

    def analyze_sweep(self):
        self.parent.analyze(self.solution_name)

    def get_network_data(self, formats):
        if isinstance(formats, str):
            formats = formats.split(",")
        formats = [f.upper() for f in formats]

        fmts_lists = {'S': [], 'Y': [], 'Z': []}

        for f in formats:
            fmts_lists[f[0]].append((int(f[1]), int(f[2])))

        ret = [None] * len(formats)
        freq = None

        for data_type, list in fmts_lists.items():
            if list:
                fn = tempfile.mktemp()
                self.parent._solutions.ExportNetworkData(
                    [],  self.parent.name + " : " + self.name,
                    2, fn, ["all"], False, 0,
                    data_type, -1, 1, 15
                )
                with open(fn) as f:
                    f.readline()
                    colnames = f.readline().split()
                array = np.loadtxt(fn, skiprows=2)
                # WARNING for python 3 probably need to use genfromtxt
                if freq is None:
                    freq = array[:, 0]
                for i, j in list:
                    real_idx = colnames.index(
                        "%s[%d,%d]_Real" % (data_type, i, j))
                    imag_idx = colnames.index(
                        "%s[%d,%d]_Imag" % (data_type, i, j))
                    c_arr = array[:, real_idx] + 1j*array[:, imag_idx]
                    ret[formats.index("%s%d%d" % (data_type, i, j))] = c_arr

        return freq, ret

    def create_report(self, name=None, expr="dB(S(port_2,port_1))"):
        if name==None:
            name=self.name
        existing = self.parent._reporter.GetAllReportNames()
        name = increment_name(name, existing)
        var_names = self.parent.parent.get_variable_names()
        var_args = sum([["%s:=" % v_name, ["Nominal"]]
                        for v_name in var_names], [])
        self.parent._reporter.CreateReport(
            name, "Modal Solution Data", "Rectangular Plot",
            self.solution_name, ["Domain:=", "Sweep"], [
                "Freq:=", ["All"]] + var_args,
            ["X Component:=", "Freq", "Y Component:=", [expr]], [])
        return HfssReport(self.parent.parent, name)

    def get_report_arrays(self, expr):
        r = self.create_report(name="Temp", expr=expr)
        return r.get_arrays()


class HfssReport(COMWrapper):
    def __init__(self, design, name):
        """
        :type design: HfssDesign
        :type name: str
        """
        super(HfssReport, self).__init__()
        self.parent_design = design
        self.name = name

    def export_to_file(self, filename):
        filepath = os.path.abspath(filename)
        self.parent_design._reporter.ExportToFile(self.name, filepath)

    def get_arrays(self):
        fn = tempfile.mktemp(suffix=".csv")
        self.export_to_file(fn)
        return np.loadtxt(fn, skiprows=1, delimiter=',').transpose()
        # warning for python 3 probably need to use genfromtxt


class Optimetrics(COMWrapper):
    """
    Optimetrics script commands executed by the "Optimetrics" module.

    Example use:
    .. code-block python
            opti = Optimetrics(pinfo.design)
            names = opti.get_setup_names()
            print('Names of optimetrics: ', names)
            opti.solve_setup(names[0])

    Note that running optimetrics requires the license for Optimetrics by Ansys.
    """

    def __init__(self, design):
        super(Optimetrics, self).__init__()

        self.design = design  # parent
        self._optimetrics = self.design._optimetrics  # <COMObject GetModule>
        self.setup_names = None

    def get_setup_names(self):
        """
        Return list of Optimetrics setup names
        """
        self.setup_names = list(self._optimetrics.GetSetupNames())
        return self.setup_names.copy()

    def solve_setup(self, setup_name: str):
        """
        Solves the specified Optimetrics setup.
        Corresponds to:  Right-click the setup in the project tree, and then click
        Analyze on the shortcut menu.

        setup_name (str) : name of setup, should be in get_setup_names

        Blocks execution until ready to use.

        Note that this requires the license for Optimetrics by Ansys.
        """
        return self._optimetrics.SolveSetup(setup_name)

    def create_setup(self, variable, swp_params, name="ParametricSetup1", swp_type='linear_step',
                     setup_name=None,
                     save_fields=True, copy_mesh=True, solve_with_copied_mesh_only=True,
                     setup_type='parametric'
                     ):
        """
        Inserts a new parametric setup.


        For  type_='linear_step' swp_params is start, stop, step:
             swp_params = ("12.8nH" "13.6nH", "0.2nH")

        Corresponds to ui access:
            Right-click the Optimetrics folder in the project tree,
            and then click Add> Parametric on the shortcut menu.
        """
        setup_name = setup_name or self.design.get_setup_names()[0]
        print(
            f"Inserting optimetrics setup `{name}` for simulation setup: `{setup_name}`")

        if setup_type != 'parametric':
            raise NotImplementedError()

        if swp_type == 'linear_step':
            assert len(swp_params) == 3
            # e.g., "LIN 12.8nH 13.6nH 0.2nH"
            swp_str = f"LIN {swp_params[0]} {swp_params[1]} {swp_params[2]}"
        else:
            raise NotImplementedError()

        self._optimetrics.InsertSetup("OptiParametric",
                                      [
                                          f"NAME:{name}",
                                          "IsEnabled:="		, True,
                                          [
                                              "NAME:ProdOptiSetupDataV2",
                                              "SaveFields:="		, save_fields,
                                              "CopyMesh:="		, copy_mesh,
                                              "SolveWithCopiedMeshOnly:=", solve_with_copied_mesh_only,
                                          ],
                                          [
                                              "NAME:StartingPoint"
                                          ],
                                          "Sim. Setups:="		, [setup_name],
                                          [
                                              "NAME:Sweeps",
                                              [
                                                  "NAME:SweepDefinition",
                                                  "Variable:="		, variable,
                                                  "Data:="		, swp_str,
                                                  "OffsetF1:="		, False,
                                                  "Synchronize:="		, 0
                                              ]
                                          ],
                                          [
                                              "NAME:Sweep Operations"
                                          ],
                                          [
                                              "NAME:Goals"
                                          ]
                                      ])
        return setup_name

class HfssModeler(COMWrapper):
    def __init__(self, design):
        """
        :type design: HfssDesign
        """
        super(HfssModeler, self).__init__()
        self.parent = design
        self._modeler = design._modeler
        self._boundaries = design._boundaries
        self._mesh = design._mesh  # Mesh module

    def set_units(self, units, rescale=True):
        self._modeler.SetModelUnits(
            ["NAME:Units Parameter", "Units:=", units, "Rescale:=", rescale])

    def get_units(self):
        """Get the model units.
            Return Value:    A string contains current model units. """
        return str(self._modeler.GetModelUnits())

    def get_all_properties(self, obj_name, PropTab='Geometry3DAttributeTab'):
        '''
            Get all properties for modeler PropTab, PropServer
        '''
        PropServer = obj_name
        properties = {}
        for key in self._modeler.GetProperties(PropTab, PropServer):
            properties[key] = self._modeler.GetPropertyValue(
                PropTab, PropServer, key)
        return properties

    def _attributes_array(self,
                          name=None,
                          nonmodel=False,
                          wireframe=False,
                          color=None,
                          transparency=0.9,
                          material=None,  # str
                          solve_inside=None,  # bool
                          coordinate_system="Global"):
        arr = ["NAME:Attributes", "PartCoordinateSystem:=", coordinate_system]
        if name is not None:
            arr.extend(["Name:=", name])

        if nonmodel or wireframe:
            flags = 'NonModel' if nonmodel else ''  # can be done smarter
            if wireframe:
                flags += '#' if len(flags) > 0 else ''
                flags += 'Wireframe'
            arr.extend(["Flags:=", flags])

        if color is not None:
            arr.extend(["Color:=", "(%d %d %d)" % color])
        if transparency is not None:
            arr.extend(["Transparency:=", transparency])
        if material is not None:
            arr.extend(["MaterialName:=", material])
        if solve_inside is not None:
            arr.extend(["SolveInside:=", solve_inside])

        return arr

    def _selections_array(self, *names):
        return ["NAME:Selections", "Selections:=", ",".join(names)]

    def assign_material(self, objs, material="copper", solve_inside=False):
        if type(objs)!=list:
            objs=[objs]
        self._modeler.AssignMaterial(["NAME:Selections",
                "AllowRegionDependentPartSelectionForPMLCreation:=", True,
                "AllowRegionSelectionForPMLCreation:=", True,
                "Selections:=",','.join(objs)
                ],
                [   
                "NAME:Attributes",
                "MaterialValue:="	, "\"%s\""%material,
                "SolveInside:="		, solve_inside,
                "IsMaterialEditable:="	, True,
                "UseMaterialAppearance:=", False,
                "IsLightweight:="	, False
                ])

    def mesh_length(self, name_mesh, objects: list, max_length='0.1mm', **kwargs):
        '''
        "RefineInside:="	, False,
        "Enabled:="		, True,
        "RestrictElem:="	, False,
        "NumMaxElem:="		, "1000",
        "RestrictLength:="	, True,
        "MaxLength:="		, "0.1mm"

        Example use:
        modeler.assign_mesh_length('mesh2', ["Q1_mesh"], MaxLength=0.1)
        '''
        assert isinstance(objects, list)

        arr = [f"NAME:{name_mesh}",
               "Objects:=", objects,
               'MaxLength:=', max_length]
        ops = ['RefineInside', 'Enabled', 'RestrictElem',
               'NumMaxElem', 'RestrictLength']
        for key, val in kwargs.items():
            if key in ops:
                if type(val)==bool:
                    arr += [key+':=', val]
                else:
                    arr += [key+':=', str(val)]
            else:
                logger.error('KEY `{key}` NOT IN ops!')

        self._mesh.AssignLengthOp(arr)

    def mesh_reassign(self, name_mesh, objects: list):
        assert isinstance(objects, list)
        self._mesh.ReassignOp(name_mesh, ["Objects:=", objects])

    def mesh_get_names(self, kind="Length Based"):
        ''' "Length Based", "Skin Depth Based", ...'''
        return list(self._mesh.GetOperationNames(kind))

    def mesh_get_all_props(self, mesh_name):
        # TODO: make mesh tis own  class with preperties
        prop_tab = 'MeshSetupTab'
        prop_server = f'MeshSetup:{mesh_name}'
        prop_names = self.parent._design.GetProperties(
            'MeshSetupTab', prop_server)
        dic = {}
        for name in prop_names:
            dic[name] = self._modeler.GetPropertyValue(
                prop_tab, prop_server, name)
        return dic

    def draw_box_corner(self, pos, size,**kwargs):
        name = self._modeler.CreateBox(
            ["NAME:BoxParameters",
             "XPosition:=", str(pos[0]),
             "YPosition:=", str(pos[1]),
             "ZPosition:=", str(pos[2]),
             "XSize:=", str(size[0]),
             "YSize:=", str(size[1]),
             "ZSize:=", str(size[2])],
            self._attributes_array(**kwargs)
        )
        return Box(name, self, pos, size)

    def draw_box_center(self, pos, size, **kwargs):
        corner_pos = [var(p) - var(s)/2 for p, s in zip(pos, size)]
        return self.draw_box_corner(corner_pos, size, **kwargs)

    def draw_polyline(self, points, closed=True, **kwargs):
        """
            Draws a closed or open polyline.
            If closed = True, then will make into a sheet.
            points : need to be in the correct units
        """
        pointsStr = ["NAME:PolylinePoints"]
        indexsStr = ["NAME:PolylineSegments"]
        for ii, point in enumerate(points):
            pointsStr.append(["NAME:PLPoint",
                              "X:=", str(point[0]),
                              "Y:=", str(point[1]),
                              "Z:=", str(point[2])])
            indexsStr.append(["NAME:PLSegment", "SegmentType:=",
                              "Line", "StartIndex:=", ii, "NoOfPoints:=", 2])
        if closed:
            pointsStr.append(["NAME:PLPoint",
                              "X:=", str(points[0][0]),
                              "Y:=", str(points[0][1]),
                              "Z:=", str(points[0][2])])
            params_closed = ["IsPolylineCovered:=",
                             True, "IsPolylineClosed:=", True]
        else:
            indexsStr = indexsStr[:-1]
            params_closed = ["IsPolylineCovered:=",
                             True, "IsPolylineClosed:=", False]

        name = self._modeler.CreatePolyline(
            ["NAME:PolylineParameters",
             *params_closed,
             pointsStr,
             indexsStr],
            self._attributes_array(**kwargs)
        )

        if closed:
            return Polyline(name, self, points)
        else:
            return OpenPolyline(name, self, points)

    def draw_rect_corner(self, pos, x_size=0, y_size=0, z_size=0, **kwargs):
        size = [x_size, y_size, z_size]
        assert 0 in size
        axis = "XYZ"[size.index(0)]
        w_idx, h_idx = {
            'X': (1, 2),
            'Y': (2, 0),
            'Z': (0, 1)
        }[axis]

        name = self._modeler.CreateRectangle(
            ["NAME:RectangleParameters",
             "XStart:=", str(pos[0]),
             "YStart:=", str(pos[1]),
             "ZStart:=", str(pos[2]),
             "Width:=",  str(size[w_idx]),
             "Height:=", str(size[h_idx]),
             "WhichAxis:=", axis],
            self._attributes_array(**kwargs)
        )
        return Rect(name, self, pos, size)

    def draw_rect_center(self, pos, x_size=0, y_size=0, z_size=0, **kwargs):
        corner_pos = [var(p) - var(s)/2. for p,
                      s in zip(pos, [x_size, y_size, z_size])]
        return self.draw_rect_corner(corner_pos, x_size, y_size, z_size,  **kwargs)

    def draw_cylinder(self, pos, radius, height, axis, **kwargs):
        assert axis in "XYZ"
        unit_suffix=self.get_units()
        return self._modeler.CreateCylinder(
            ["NAME:CylinderParameters",
             "XCenter:=", pos[0],
             "YCenter:=", pos[1],
             "ZCenter:=", pos[2],
             "Radius:=", radius,
             "Height:=", height,
             "WhichAxis:=", axis,
             "NumSides:=", 0],
            self._attributes_array(**kwargs))

    def draw_cylinder_center(self, pos, radius, height, axis, **kwargs):
        axis_idx = ["X", "Y", "Z"].index(axis)
        edge_pos = copy(pos)
        edge_pos[axis_idx] = var(pos[axis_idx]) - var(height)/2
        return self.draw_cylinder(edge_pos, radius, height, axis, **kwargs)


    def draw_wirebond(self, pos, ori, width, height='0.1mm', z=0,wire_diameter="0.02mm", NumSides=6,**kwargs):
        '''
            Args:
                pos: 2D positon vector  (specify center point)
                ori: should be normed
                z: z postion

            # TODO create Wirebond class
            psoition is the origin of one point
            ori is the orientation vector, which gets normalized
        '''
        p = np.array(self.parent.conv_variable_value(pos))
        o = np.array(self.parent.conv_variable_value(ori))
        ori=self.parent.conv_variable_value(ori)
        pad1 = p-o*self.parent.conv_variable_value(width/2.)
        name = self._modeler.CreateBondwire(["NAME:BondwireParameters",
                                            "WireType:=", "Low",
                                            "WireDiameter:=", wire_diameter,
                                            "NumSides:=", NumSides,
                                            "XPadPos:=", str(pad1[0]),
                                            "YPadPos:=", pad1[1],
                                            "ZPadPos:=", z,
                                            "XDir:=", ori[0],
                                            "YDir:=", ori[1],
                                            "ZDir:=", 0,
                                            "Distance:=", width,
                                            "h1:=", height,
                                            "h2:=", "0mm",
                                            "alpha:=", "80deg",
                                            "beta:=", "80deg",
                                            "WhichAxis:=", "Z"],
                                            self._attributes_array(**kwargs))

        return name

    def draw_region(self, Padding, PaddingType="Percentage Offset", name='Region',
                    material="\"vacuum\""):
        """
            PaddingType : 'Absolute Offset', "Percentage Offset"
        """
        # TODO: Add option to modify these
        RegionAttributes = [
            "NAME:Attributes",
            "Name:="		, name,
            "Flags:="		, "Wireframe#",
            "Color:="		, "(255 0 0)",
            "Transparency:="	, 1,
            "PartCoordinateSystem:=", "Global",
            "UDMId:="		, "",
            "IsAlwaysHiden:="	, False,
            "MaterialValue:="	, material,
            "SolveInside:="		, True
        ]

        self._modeler.CreateRegion(
            [
                "NAME:RegionParameters",
                "+XPaddingType:="	, PaddingType,
                "+XPadding:="		, Padding[0][0],
                "-XPaddingType:="	, PaddingType,
                "-XPadding:="		, Padding[0][1],
                "+YPaddingType:="	, PaddingType,
                "+YPadding:="		, Padding[1][0],
                "-YPaddingType:="	, PaddingType,
                "-YPadding:="		, Padding[1][1],
                "+ZPaddingType:="	, PaddingType,
                "+ZPadding:="		, Padding[2][0],
                "-ZPaddingType:="	, PaddingType,
                "-ZPadding:="		, Padding[2][1]
            ],
            RegionAttributes)

    def unite(self, names, keep_originals=False):
        self._modeler.Unite(
            self._selections_array(*names),
            ["NAME:UniteParameters", "KeepOriginals:=", keep_originals]
        )
        return names[0]

    def intersect(self, names, keep_originals=False):
        self._modeler.Intersect(
            self._selections_array(*names),
            ["NAME:IntersectParameters", "KeepOriginals:=", keep_originals]
        )
        return names[0]

    def separate(self, names, split_plane='XY', keep_originals='positive'):
        assert keep_originals in ['positive', 'negative', 'both'], "keep originals must be positive, negative, or both"
        assert split_plane in ['XY', "ZX", 'YZ'], "Split plane must be in XY, ZX, YZ plane"
        if keep_originals=='positive' or keep_originals=='negative':
            which_side=keep_originals.capitalize()+'Only'
        else:
            which_side=keep_originals.capitalize()
        self._modeler.Split([
                                "NAME:Selections",
                                "Selections:="		, names,
                                "NewPartsModelFlag:="	, "Model"
                            ], 
                            [
                                "NAME:SplitToParameters",
                                "SplitPlane:="		, split_plane,
                                "WhichSide:="		, which_side,
                                "ToolType:="		, "PlaneTool",
                                "ToolEntityID:="	, -1,
                                "SplitCrossingObjectsOnly:=", False,
                                "DeleteInvalidObjects:=", True
                            ])
        return names

    def translate(self, name, vector):
        self._modeler.Move(
            self._selections_array(name),
            ["NAME:TranslateParameters",
             "TranslateVectorX:=", vector[0],
             "TranslateVectorY:=", vector[1],
             "TranslateVectorZ:=", vector[2]]
        )

    def copy_object(self, copy_obj):
        self._modeler.Copy([
		                    "NAME:Selections",
		                    "Selections:="		, copy_obj
	                        ])

    def paste_object(self):
        try:
            return self._modeler.Paste()[0]
        except:
            raise Exception('No object to paste')

    def create_objects_from_faces(self, obj, faces, make_entity=True):
        if type(faces)!=list:
            faces=[faces]

        detached_faces=self._modeler.CreateObjectFromFaces(
                        [
                            "NAME:Selections",
                            "Selections:="		, obj,
                            "NewPartsModelFlag:="	, "Model"
                        ],
                        [
                            "NAME:Parameters",
                            [
                                "NAME:BodyFromFaceToParameters",
                                "FacesToDetach:="	, faces
                            ]
                            ],
                            [
                            "CreateGroupsForNewObjects:=", False
                            ]
                        )
        if make_entity==True:
            if len(detached_faces)==1:
                return Face(detached_faces[0], self)
            else:
                face_coll=[]
                for face in list(detached_faces):
                    face_coll.append(Face(face, self))
                return face_coll
        else:
            if len(detached_faces)==1:
                return detached_faces[0]
            else:
                return list(detached_faces)


    def get_boundary_assignment(self, boundary_name: str):
        # Gets a list of face IDs associated with the given boundary or excitation assignment.
        objects = self._boundaries.GetBoundaryAssignment(boundary_name)
        # Gets an object name corresponding to the input face id. Returns the name of the corresponding object name.
        objects = [self._modeler.GetObjectNameByFaceID(k) for k in objects]
        return objects


    def append_PerfE_assignment(self, boundary_name: str, object_names: list):
        '''
            This will create a new boundary if need, and will
            otherwise append given names to an exisiting boundary
        '''
        # enforce
        boundary_name = str(boundary_name)
        if isinstance(object_names, str):
            object_names = [object_names]
        object_names = list(object_names)  # enforce list

        # do actual work
        if boundary_name not in self._boundaries.GetBoundaries():  # GetBoundariesOfType("Perfect E")
            # need to make a new boundary
            self.assign_perfect_E(object_names, name=boundary_name)
        else:
            # need to append
            objects = list(self.get_boundary_assignment(boundary_name))
            self._boundaries.ReassignBoundary(["NAME:" + boundary_name,
                                               "Objects:=", list(set(objects + object_names))])

    def append_mesh(self, mesh_name: str, object_names: list, old_objs: list,
                    **kwargs):
        '''
        This will create a new boundary if need, and will
        otherwise append given names to an exisiting boundary
        old_obj = circ._mesh_assign
        '''
        mesh_name = str(mesh_name)
        if isinstance(object_names, str):
            object_names = [object_names]
        object_names = list(object_names)  # enforce list

        if mesh_name not in self.mesh_get_names():  # need to make a new boundary
            objs = object_names
            self.mesh_length(mesh_name, object_names, **kwargs)
        else:  # need to append
            objs = list(set(old_objs + object_names))
            self.mesh_reassign(mesh_name, objs)

        return objs

    def assign_thin_conductor(self, obj, name='ThinCond', material='Copper', thickness="50nm", direction='Positive'):
        if not isinstance(obj, list):
            obj = [obj]
        params= ["NAME:"+name, 
                "Objects:=", obj,
                "Material:=", material, 
                "Thickness:=", thickness, 
                "Direction:=",  direction]
        try:
            surf=self._boundaries.AssignThinConductor(params)
        except:
            desktop=self.parent.parent.parent
            project_name=self.parent.parent.name
            design_name=self.parent.name
            ERR=self.parent.parent.parent.GetMessages(project_name, design_name, 1)
            for errors in ERR:
                if 'Duplicate  name: \'%s\''%name in errs:
                    name=increment_name(name, [name])
                    params[0]="NAME:"+name
            surf=self._boundaries.AssignThinConductor(params)
        return surf

    def assign_net(self, obj, name='Net'):
        if not isinstance(obj, list):
            obj = [obj]
        params=['NAME:'+name, 
                'Objects:=', obj
                ]
        return self._boundaries.AssignNet(params)


    def assign_perfect_E(self, obj, face=None, name='PerfE'):
        '''
            Takes a name of an object or a list of object names.
            If `name` is not specified `PerfE` is appended to object name for the name.
        '''
        if face==None:
            if not isinstance(obj, list):
                obj = [obj]
            if name == 'PerfE':
                name = str(obj[-1])+'_'+name
            name = increment_name(name, self._boundaries.GetBoundaries())
            self._boundaries.AssignPerfectE(
                ["NAME:"+name, "Objects:=", obj, "InfGroundPlane:=", False])
        else:
            if not isinstance(face, list):
                face=[face]
            name = increment_name(name, self._boundaries.GetBoundaries())   
            self._boundaries.AssignPerfectE(
                ["NAME:"+name, "Faces:=", face, "InfGroundPlane:=", False])

    def assign_impedance(self, res, reac, obj, face=None, name='Imped'):
        '''
            Takes a name of an object or a list of object names.
            If `name` is not specified `PerfE` is appended to object name for the name.
        '''
        if face==None:
            if not isinstance(obj, list):
                obj = [obj]
            if name == 'Imped':
                name = str(obj[-1])+'_'+name
            name = increment_name(name, self._boundaries.GetBoundaries())
            params=["NAME:"+name,
                "Resistance:=", str(res),
                "Reactance:=", str(reac),
                "Objects:=", obj,
                "InfGroundPlane:=", False
                ]
        else:
            if not isinstance(face, list):
                face=[face]
            name = increment_name(name, self._boundaries.GetBoundaries())  
            params=["NAME:"+name,
                "Resistance:=", str(res),
                "Reactance:=", str(reac),
                "Faces:=", face,
                "InfGroundPlane:=", False
                ]
        self._boundaries.AssignImpedance(params)

    def assign_finite_conductivity(self, face, material=None, name="FiniteCond", params=None, units='um'):
        if name=="FiniteCond":
            name = increment_name(name, self._boundaries.GetBoundaries())
        else:
            pass

        if self.parent.parent.get_material_props(material)==[]:
            raise Exception('ERROR: %s not a valid project material'%material)
        else:
            pass

        if type(face)==list:
            pass
        else:
            face=[face]

        default_props={
            "UseMaterial":True,
            "Material":material,
            "Conductivity":None,
            "Permeability":None,
            "Roughness":"2um",
            "InfGroundPlane":False,
            "Objects":face,
            "Radius":1,
            "Ratio":1
            }

        if params==None:
            params=default_props
        else:
            pass

        props_dict={}

        if type(params)==list:
            for key in list(params)[::2]:
                if key in default_props:
                    val=list(params)[params.index(key)+1]
                    if type(val)==float or type(val)==int:
                        if key=="Roughness" or key=="Radius":
                            props_dict[key]=str(val)+units
                        else:
                            props_dict[key]=str(val)
                    else:
                        props_dict[key]=val
                else:
                    print('%s parameter not valid'%key)
        elif type(params)==dict:
            for key in iter(params.keys()):
                if key in default_props:
                    val=params[key]
                    if type(val)==float or type(val)==int:
                        if key=="Roughness" or key=="Radius":
                            props_dict[key]=str(val)+units
                        else:
                            props_dict[key]=str(val)
                    else:
                        props_dict[key]=val
                else:
                    print('%s parameter not valid'%key)
        else:
            raise Exception('ERROR: Params must be default property list or dict object.')

        params=["NAME:"+name]
        for key in iter(props_dict.keys()):
            if props_dict[key]!=None: 
                params.append(key+":=")
                params.append(props_dict[key])
            else:
                pass
            
        self._boundaries.AssignFiniteCond(params)

    def _make_lumped_rlc(self, r, l, c, start, end, obj_arr, name="LumpRLC"):
        name = increment_name(name, self._boundaries.GetBoundaries())
        params = ["NAME:"+name]
        params += obj_arr
        params.append(["NAME:CurrentLine",
                       # for some reason here it seems to swtich to use the model units, rather than meters
                       "Start:=", fix_units(start, unit_assumed=LENGTH_UNIT),
                       "End:=",   fix_units(end, unit_assumed=LENGTH_UNIT)])
        params += ["UseResist:=", r != 0, "Resistance:=", r,
                   "UseInduct:=", l != 0, "Inductance:=", l,
                   "UseCap:=", c != 0, "Capacitance:=", c]
        self._boundaries.AssignLumpedRLC(params)

    def _make_lumped_port(self, start, end, obj_arr, z0="50ohm", name="LumpPort"):
        start = fix_units(start, unit_assumed=LENGTH_UNIT)
        end = fix_units(end, unit_assumed=LENGTH_UNIT)

        name = increment_name(name, self._boundaries.GetBoundaries())
        params = ["NAME:"+name]
        params += obj_arr
        params += ["RenormalizeAllTerminals:=", True, "DoDeembed:=", False,
                   ["NAME:Modes", ["NAME:Mode1",
                                   "ModeNum:=", 1,
                                   "UseIntLine:=", True,
                                   ["NAME:IntLine",
                                    "Start:=", start,
                                    "End:=",   end],
                                   "CharImp:=", "Zpi",
                                   "AlignmentGroup:=", 0,
                                   "RenormImp:=", "50ohm"]],
                   "ShowReporterFilter:=", False, "ReporterFilter:=", [True],
                   "FullResistance:=", "50ohm", "FullReactance:=", "0ohm"]

        self._boundaries.AssignLumpedPort(params)

    def get_face_ids(self, obj):
        return self._modeler.GetFaceIDs(obj)

    def get_face_id_by_pos(self, obj, pos):
        x_pos=pos[0]
        y_pos=pos[1]
        z_pos=pos[2]
        params=["NAME:FaceParameters",
                "BodyName:=", obj, 
                "XPosition:=", str(x_pos), 
                "YPosition:=", str(y_pos), 
                "ZPosition:=", str(z_pos)]
        try:
            return self._modeler.GetFaceByPosition(params)
        except:
            print('No face found at position %.2s, %.2s, %.2s'%(str(x_pos), str(y_pos), str(z_pos)))
            return None

    def get_edge_ids_by_face(self, face):
        return list(self._modeler.GetEdgeIDsFromFace(int(face)))


    def get_object_name_by_face_id(self, ID: str):
        ''' Gets an object name corresponding to the input face id. '''
        return self._modeler.GetObjectNameByFaceID(ID)

    def get_object_names(self):
        i=0
        names=[]
        while True:
            try:
                names.append(self._modeler.GetObjectName(i))
            except:
                break
            i+=1
        return names

    def get_vertex_ids(self, obj):
        """
            Get the vertex IDs of given an object name
            oVertexIDs = oEditor.GetVertexIDsFromObject(“Box1”)
        """
        return self._modeler.GetVertexIDsFromObject(obj)

    def eval_expr(self, expr, units="mm"):
        if not isinstance(expr, str):
            return expr
        return self.parent.eval_expr(expr, units)

    def get_objects_in_group(self, group):
        """
        Use:              Returns the objects for the specified group.
        Return Value:    The objects in the group.
        Parameters:      <groupName>  Type: <string>
        One of  <materialName>, <assignmentName>, "Non Model",
                "Solids", "Unclassi­fied", "Sheets", "Lines"
        """
        return list(self._modeler.GetObjectsInGroup(group))


    def set_working_coordinate_system(self, cs_name="Global"):
        """
        Use:                   Sets the working coordinate system.
        Command:         Modeler>Coordinate System>Set Working CS
        """
        self._modeler.SetWCS(
            [
                "NAME:SetWCS Parameter",
                "Working Coordinate System:=", cs_name,
                "RegionDepCSOk:="	, False  # this one is prob not needed, but comes with the record tool
            ])

    def create_relative_coorinate_system_both(self, cs_name,
                                              origin=["0um", "0um", "0um"],
                                              XAxisVec=["1um", "0um", "0um"],
                                              YAxisVec=["0um", "1um", "0um"]):
        """
        Use:     Creates a relative coordinate system. Only the    Name attribute of the <AttributesArray> parameter is supported.
        Command: Modeler>Coordinate System>Create>Relative CS->Offset
        Modeler>Coordinate System>Create>Relative CS->Rotated
        Modeler>Coordinate System>Create>Relative CS->Both

        Current cooridnate system is set right after this.

        cs_name : name of coord. sys
            If the name already exists, then a new coordinate system with _1 is created.

        origin, XAxisVec, YAxisVec: 3-vectors
            You can also pass in params such as origin = [0,1,0] rather than ["0um","1um","0um"], but these will be interpreted in default units, so it is safer to be explicit. Explicit over implicit.
        """
        self._modeler.CreateRelativeCS(
            [
                "NAME:RelativeCSParameters",
                "Mode:="		, "Axis/Position",
                "OriginX:="		, origin[0],
                "OriginY:="		, origin[1],
                "OriginZ:="		, origin[2],
                "XAxisXvec:="		, XAxisVec[0],
                "XAxisYvec:="		, XAxisVec[1],
                "XAxisZvec:="		, XAxisVec[2],
                "YAxisXvec:="		, YAxisVec[0],
                "YAxisYvec:="		, YAxisVec[1],
                "YAxisZvec:="		, YAxisVec[1]
            ],
            [
                "NAME:Attributes",
                "Name:="		, cs_name
            ])

    def subtract(self, blank_name, tool_names, keep_originals=False):
        selection_array = ["NAME:Selections",
                           "Blank Parts:=", blank_name,
                           "Tool Parts:=", ",".join(tool_names)]
        self._modeler.Subtract(
            selection_array,
            ["NAME:UniteParameters", "KeepOriginals:=", keep_originals]
        )
        return blank_name

    def rotate(self, selection_name, axis, angle):
        assert axis.lower() in "xyz", 'Axis must be X,Y,Z'
        if type(selection_name)!=list:
            selection_name=[selection_name]
        selection_array = ["NAME:Selections",
                           "Selections:="	, ",".join(selection_name),
                           "NewPartsModelFlag:="	, "Model"]
        self._modeler.Rotate(
            selection_array,
            ["NAME:RotateParameters",
              "RotateAxis:=", axis,
              "RotateAngle:=", str(angle)]
        )
        return selection_name

    def _fillet(self, radius, vertex_index, obj):
        vertices = self._modeler.GetVertexIDsFromObject(obj)
        if isinstance(vertex_index, list):
            to_fillet = [int(vertices[v]) for v in vertex_index]
        else:
            to_fillet = [int(vertices[vertex_index])]

        self._modeler.Fillet(["NAME:Selections", "Selections:=", obj],
                             ["NAME:Parameters",
                              ["NAME:FilletParameters",
                               "Edges:=", [],
                               "Vertices:=", to_fillet,
                               "Radius:=", radius,
                               "Setback:=", "0mm"]])

    def _fillet_edges(self, radius, edge_index, obj):
        edges = self._modeler.GetEdgeIDsFromObject(obj)
        if isinstance(edge_index, list):
            to_fillet = [int(edges[e]) for e in edge_index]
        else:
            to_fillet = [int(edges[edge_index])]

        self._modeler.Fillet(["NAME:Selections", "Selections:=", obj],
                             ["NAME:Parameters",
                              ["NAME:FilletParameters",
                               "Edges:=", to_fillet,
                               "Vertices:=", [],
                               "Radius:=", radius,
                               "Setback:=", "0mm"]])

    def _fillets(self, radius, vertices, obj):
        self._modeler.Fillet(["NAME:Selections", "Selections:=", obj],
                             ["NAME:Parameters",
                              ["NAME:FilletParameters",
                               "Edges:=", [],
                               "Vertices:=", vertices,
                               "Radius:=", radius,
                               "Setback:=", "0mm"]])

    def _sweep_along_path(self, to_sweep, path_obj):
        self.rename_obj(path_obj, str(path_obj)+'_path')
        new_name = self.rename_obj(to_sweep, path_obj)
        names = [path_obj, str(path_obj)+'_path']
        self._modeler.SweepAlongPath(self._selections_array(*names),
                                     ["NAME:PathSweepParameters",
                                      "DraftAngle:="		, "0deg",
                                      "DraftType:="		, "Round",
                                      "CheckFaceFaceIntersection:=", False,
                                      "TwistAngle:="		, "0deg"])
        return Polyline(new_name, self)

    def sweep_along_vector(self, names, vector):
        self._modeler.SweepAlongVector(self._selections_array(*names),
                                       ["NAME:VectorSweepParameters",
                                        "DraftAngle:="		, "0deg",
                                        "DraftType:="		, "Round",
                                        "CheckFaceFaceIntersection:=", False,
                                        "SweepVectorX:="	, vector[0],
                                        "SweepVectorY:="	, vector[1],
                                        "SweepVectorZ:="	, vector[2]
                                        ])

    def duplicate_along_vector(self, obj, vector, num_clones, params=None):
        
        unit_suffix=self.get_units()
        if type(obj)==list:
            pass
        else:
            obj=[obj]
            
        default_props={
                    "NewPartsModelFlag":"Model",
                    "CreateNewObjects":True,
                    "XComponent":str(vector[0])+unit_suffix,
                    "YComponent":str(vector[1])+unit_suffix,
                    "ZComponent":str(vector[2])+unit_suffix,
                    "NumClones":str(num_clones),
                    "DuplicateBoundaries":False
                    }

        if params==None:
            params=default_props
        else:
            pass

        props_dict={}

        if type(params)==list:
            for key in list(params)[::2]:
                if key in default_props:
                    val=list(params)[params.index(key)+1]
                    default_props[key]=val
                else:
                    print('%s parameter not valid'%key)
        elif type(params)==dict:
            for key in iter(params.keys()):
                if key in default_props:
                    val=params[key]
                    default_props[key]=val
                else:
                    print('%s parameter not valid'%key)
        else:
            raise Exception('ERROR: Params must be default property list or dict object.')
            
        for objs in obj:
            selections=["NAME:Selections", "Selections:=", objs]
            duplicate_params=["NAME:DuplicateToAlongLineParameters"]
            options=["NAME:Options"]
            for key in iter(default_props.keys()):
                if key in ["NewPartsModelFlag"]:
                    selections.append(key+":=")
                    selections.append(default_props[key])
                elif key in ["CreateNewObjects", "XComponent", "YComponent", "ZComponent", "NumClones"]:
                    duplicate_params.append(key+":=")
                    duplicate_params.append(default_props[key])
                elif key in ["DuplicateBoundaries"]:
                    options.append(key+":=")
                    options.append(default_props[key])

            self._modeler.DuplicateAlongLine(selections, duplicate_params, options)
        

    def rename_obj(self, obj, name):
        self._modeler.ChangeProperty(["NAME:AllTabs",
                                      ["NAME:Geometry3DAttributeTab",
                                       ["NAME:PropServers", str(obj)],
                                       ["NAME:ChangedProps", ["NAME:Name", "Value:=", str(name)]]]])
        return name

    def assign_non_model(self, obj):
        self._modeler.ChangeProperty(["NAME:AllTabs",
		["NAME:Geometry3DAttributeTab",
            ["NAME:PropServers", str(obj)],
			["NAME:ChangedProps",
				["NAME:Model","Value:="	, False]]]])

    def import_3D_obj(self, path):
        source=["NAME:NativeBodyParameters",
                "SourceFile:=", path]
        try:
            self._modeler.Import(source)
            return self.get_object_names()[-1]
        except:
            print('UNABLE TO IMPORT 3D OBJECT')
    
    def import_DXF(self, path, layers=None, scale=1E-6, self_stitch=True, sheet_bodies=False):
        try:
            file=ezdxf.readfile(path)
        except:
            print("Unable to open specified DXF file, check path.")
            return 
        des_layers=[]
        for layer in file.layers:
            layer_name=layer.dxf.name
            if layer_name!='PYDXF' and layer_name!='Defpoints':
                des_layers.append(layer_name)
        if layers!=None:
            if type(layers)!=list:
                layers=[layers]
            layers=[lay for lay in layers if lay in des_layers]
            if layers==[]:
                raise Exception("ERROR: No selected layers is in DXF design file provided")
        else:
            layers=des_layers

        layer_info=["NAME:LayerInfo"]
        layer_props=["NAME:TechFileLayers"]  
        colors=["purple", "blue", "red", "yellow", "green", "orange"]  
        for I, lay in enumerate(layers):
            layer_info.append([ "NAME:%s"%lay,
                                "source:="		, lay,
                                "display_source:="	, lay,
                                "import:="		, True,
                                "dest:="		, "DXF_layer_%s"%lay,
                                "dest_selected:="	, False,
                                "layer_type:="		, "signal"
                            ])
            layer_props.append( "layer:=")
            layer_props.append(["name:=", lay,
                                "color:=", colors[I%len(colors)],
                                "elev:=", 0,
                                "thick:=", "0.0000000000000001m"
                                ])
        source=["NAME:options", 
                "FileName:=",path,
                "Scale:=", scale,
                "SelfStitch:=", self_stitch,
                "UnionOverlapping:=", True,
                "AutoDetectClosed:=", True,
                "DefeatureGeometry:=", False,
                "DefeatureDistance:=", 1E-16,
                "RoundCoordinates:=", False,
                "RoundNumDigits:=", 10,
                "WritePolyWithWidthAsFilledPoly:=",False,
                "ImportMethod:=", 1,
                "2DSheetBodies:=", sheet_bodies,
                layer_info,
                layer_props
                ]
        self._modeler.ImportDXF(source)

        layer_groups=[]
        model_objs=self.get_object_names()
        for lay in layers: 
            self.create_group([obj for obj in model_objs if lay in obj])
            layer_groups.append("DXF_layer_%s_Group"%lay)
        return layer_groups

    def create_group(self, objs):
        if type(objs)!=list:
            objs=[objs]

        str_list=""
        for obj in objs:
            str_list+="%s,"%obj
        
        str_list=str_list[0:-1]

        source=["NAME:GroupParameter", 
                "ParentGroupID:=", "Model", 
                "Parts:=", str_list, 
                "SubmodelInstances:=", "", 
                "Groups:=", ""]
        self._modeler.CreateGroup(source)
        return str_list 
        
    
class ModelEntity(str, HfssPropertyObject):
    prop_tab = "Geometry3DCmdTab"
    model_command = None
    transparency = make_float_prop(
        "Transparent", prop_tab="Geometry3DAttributeTab", prop_server=lambda self: self)
    material = make_str_prop(
        "Material", prop_tab="Geometry3DAttributeTab", prop_server=lambda self: self)
    wireframe = make_float_prop(
        "Display Wireframe", prop_tab="Geometry3DAttributeTab", prop_server=lambda self: self)
    coordinate_system = make_str_prop("Coordinate System")

    def __new__(self, val, *args, **kwargs):
        return str.__new__(self, val)

    def __init__(self, val, modeler):
        """
        :type val: str
        :type modeler: HfssModeler
        """
        super(ModelEntity, self).__init__(
        )  # val) #Comment out keyword to match arguments
        self.modeler = modeler
        self.prop_server = self + ":" + self.model_command + ":1"


class Box(ModelEntity):
    model_command = "CreateBox"
    position = make_float_prop("Position")
    x_size = make_float_prop("XSize")
    y_size = make_float_prop("YSize")
    z_size = make_float_prop("ZSize")

    def __init__(self, name, modeler, corner, size):
        """
        :type name: str
        :type modeler: HfssModeler
        :type corner: [(VariableString, VariableString, VariableString)]
        :param size: [(VariableString, VariableString, VariableString)]
        """
        super(Box, self).__init__(name, modeler)
        self.modeler = modeler
        self.prop_holder = modeler._modeler

        #can now accept both variable names and values: A Oriani 

        self.size=size
        self.corner=corner
        
        self.center = [var(c) + var(s)/2 for c, s in zip(self.corner, self.size)]
        faces = modeler.get_face_ids(self)
        self.z_back_face, self.z_front_face = faces[0], faces[1]
        self.y_back_face, self.y_front_face = faces[2], faces[4]
        self.x_back_face, self.x_front_face = faces[3], faces[5]


    
class Rect(ModelEntity):
    model_command = "CreateRectangle"
    # TODO: Add a rotated rectangle object.
    # Will need to first create rect, then apply rotate operation.

    def __init__(self, name, modeler, corner, size):
        super(Rect, self).__init__(name, modeler)
        self.prop_holder = modeler._modeler
        self.modeler=modeler
        self.name=name
        #can now accept both variable names and values: A Oriani 

        self.size=size
        self.corner=corner

        self.center = [var(c) + var(s)/2 if s else c for c, s in zip(corner, size)]

    def make_center_line(self, axis):
        '''
        Returns `start` and `end` list of 3 coordinates
        '''
        axis_idx = ["x", "y", "z"].index(axis.lower())
        start = [c for c in self.center]
        start[axis_idx] -= self.size[axis_idx]/2
        start = [self.modeler.eval_expr(s) for s in start]
        end = [c for c in self.center]
        end[axis_idx] += self.size[axis_idx]/2
        end = [self.modeler.eval_expr(s) for s in end]
        return start, end

    def make_rlc_boundary(self, axis, r=0, l=0, c=0, name="LumpRLC"):
        start, end = self.make_center_line(axis)
        self.modeler._make_lumped_rlc(
            r, l, c, start, end, ["Objects:=", [self]], name=name)

    def make_lumped_port(self, axis, z0="50ohm", name="LumpPort"):
        start, end = self.make_center_line(axis)
        self.modeler._make_lumped_port(
            start, end, ["Objects:=", [self]], z0=z0, name=name)

    def make_thin_conductor(self, name=None, material='Copper', thickness='50nm', direction='positive'):
        if name==None:
            name=self.name+'_thin_cond'
        self.modeler.assign_thin_conductor(self, name, material, thickness, direction)

    def make_finite_conductivity(self, material="copper", inf_ground_plane=False):
        self.modeler._boundaries.AssignFiniteCond(	[
                                                    "NAME:%s"%self.name+"_finite_cond",
                                                    "Objects:="		, [self.name],
                                                    "UseMaterial:="		, True,
                                                    "Material:="		, material,
                                                    "UseThickness:="	, False,
                                                    "Roughness:="		, "0um",
                                                    "InfGroundPlane:="	, False,
                                                    "IsTwoSided:="		, False,
                                                    "IsInternal:="		, True
                                                ])

    def make_net(self, name=None):
        if name==None:
            name=self.name+'_Net'
        self.modeler.assign_net(self, name)


class Face(ModelEntity):
    model_command='DetachFaces'
    def __init__(self, name, modeler):
        super(Face, self).__init__(name, modeler)
        self.prop_holder = modeler._modeler
        self.modeler=modeler
        self.name=name

    def make_rlc_boundary(self, axis, r=0, l=0, c=0, name="LumpRLC"):
        start, end = self.make_center_line(axis)
        self.modeler._make_lumped_rlc(
            r, l, c, start, end, ["Objects:=", [self]], name=name)

    def make_lumped_port(self, axis, z0="50ohm", name="LumpPort"):
        start, end = self.make_center_line(axis)
        self.modeler._make_lumped_port(
            start, end, ["Objects:=", [self]], z0=z0, name=name)

    def make_thin_conductor(self, name=None, material='Copper', thickness='50nm', direction='positive'):
        if name==None:
            name=self.name+'_thin_cond'
        self.modeler.assign_thin_conductor(self, name, material, thickness, direction)

    def make_finite_conductivity(self, material="copper", inf_ground_plane=False):
        self.modeler._boundaries.AssignFiniteCond(	[
                                                    "NAME:%s"%self.name+"_finite_cond",
                                                    "Objects:="		, [self.name],
                                                    "UseMaterial:="		, True,
                                                    "Material:="		, material,
                                                    "UseThickness:="	, False,
                                                    "Roughness:="		, "0um",
                                                    "InfGroundPlane:="	, False,
                                                    "IsTwoSided:="		, False,
                                                    "IsInternal:="		, True
                                                ])

    def make_net(self, name=None):
        if name==None:
            name=self.name+'_Net'
        self.modeler.assign_net(self, name)


class Polyline(ModelEntity):
    '''
        Assume closed polyline, which creates a polygon.
    '''

    model_command = "CreatePolyline"

    def __init__(self, name, modeler, points=None):
        super(Polyline, self).__init__(name, modeler)
        self.prop_holder = modeler._modeler
        if points is not None:
            self.points = points
            self.n_points = len(points)
        else:
            pass
            # TODO: points = collection of points
#        axis = find_orth_axis()

# TODO: find the plane of the polyline for now, assume Z
#    def find_orth_axis():
#        X, Y, Z = (True, True, True)
#        for point in points:
#            X =

    def unite(self, list_other):
        union = self.modeler.unite(self + list_other)
        return Polyline(union, self.modeler)

    def make_center_line(self, axis):  # Expects to act on a rectangle...
        # first : find center and size
        center = [0, 0, 0]

        for point in self.points:
            center = [center[0]+point[0]/self.n_points,
                      center[1]+point[1]/self.n_points,
                      center[2]+point[2]/self.n_points]
        size = [2*(center[0]-self.points[0][0]),
                2*(center[1]-self.points[0][1]),
                2*(center[1]-self.points[0][2])]
        axis_idx = ["x", "y", "z"].index(axis.lower())
        start = [c for c in center]
        start[axis_idx] -= size[axis_idx]/2
        start = [self.modeler.eval_var_str(
            s, unit=LENGTH_UNIT) for s in start]  # TODO
        end = [c for c in center]
        end[axis_idx] += size[axis_idx]/2
        end = [self.modeler.eval_var_str(s, unit=LENGTH_UNIT) for s in end]
        return start, end

    def make_rlc_boundary(self, axis, r=0, l=0, c=0, name="LumpRLC"):
        name = str(self)+'_'+name
        start, end = self.make_center_line(axis)
        self.modeler._make_lumped_rlc(
            r, l, c, start, end, ["Objects:=", [self]], name=name)

    def fillet(self, radius, vertex_index):
        self.modeler._fillet(radius, vertex_index, self)

    def vertices(self):
        return self.modeler.get_vertex_ids(self)

    def rename(self, new_name):
        '''
            Warning: The  increment_name only works if the sheet has not been stracted or used as a tool elsewher.
            These names are not checked - They require modifying get_objects_in_group

        '''
        new_name = increment_name(new_name, self.modeler.get_objects_in_group(
            "Sheets"))  # this is for a clsoed polyline

        # check to get the actual new name in case there was a suibtracted ibjet with that namae
        face_ids = self.modeler.get_face_ids(str(self))
        self.modeler.rename_obj(self, new_name)  # now rename
        if len(face_ids) > 0:
            new_name = self.modeler.get_object_name_by_face_id(face_ids[0])
        return Polyline(str(new_name), self.modeler)


class OpenPolyline(ModelEntity):  # Assume closed polyline
    model_command = "CreatePolyline"
    show_direction = make_prop(
        'Show Direction', prop_tab="Geometry3DAttributeTab", prop_server=lambda self: self)

    def __init__(self, name, modeler, points=None):
        super(OpenPolyline, self).__init__(name, modeler)
        self.prop_holder = modeler._modeler
        if points is not None:
            self.points = points
            self.n_points = len(points)
        else:
            pass
#        axis = find_orth_axis()

# TODO: find the plane of the polyline for now, assume Z
#    def find_orth_axis():
#        X, Y, Z = (True, True, True)
#        for point in points:
#            X =
    def vertices(self):
        return self.modeler.get_vertex_ids(self)

    def fillet(self, radius, vertex_index):
        self.modeler._fillet(radius, vertex_index, self)

    def fillets(self, radius, do_not_fillet=[]):
        '''
            do_not_fillet : Index list of verteces to not fillete
        '''
        raw_list_vertices = self.modeler.get_vertex_ids(self)
        list_vertices = []
        for vertex in raw_list_vertices[1:-1]:  # ignore the start and finish
            list_vertices.append(int(vertex))
        list_vertices = list(map(int, np.delete(list_vertices,
                                                np.array(do_not_fillet, dtype=int)-1)))
        #print(list_vertices, type(list_vertices[0]))
        if len(list_vertices) != 0:
            self.modeler._fillets(radius, list_vertices, self)
        else:
            pass

    def sweep_along_path(self, to_sweep):
        return self.modeler._sweep_along_path(to_sweep, self)

    def rename(self, new_name):
        '''
            Warning: The  increment_name only works if the sheet has not been stracted or used as a tool elsewher.
            These names are not checked - They require modifying get_objects_in_group
        '''
        new_name = increment_name(
            new_name, self.modeler.get_objects_in_group("Lines"))
        # , self.points)
        return OpenPolyline(self.modeler.rename_obj(self, new_name), self.modeler)

    def copy(self, new_name):
        new_obj = OpenPolyline(self.modeler.copy(self), self.modeler)
        return new_obj.rename(new_name)


class HfssFieldsCalc(COMWrapper):
    def __init__(self, setup):
        """
        :type setup: HfssSetup
        """
        self.setup = setup
        super(HfssFieldsCalc, self).__init__()
        self.parent = setup
        self.Mag_E = NamedCalcObject("Mag_E", setup)
        self.Mag_H = NamedCalcObject("Mag_H", setup)
        self.Mag_Jsurf = NamedCalcObject("Mag_Jsurf", setup)
        self.Mag_Jvol = NamedCalcObject("Mag_Jvol", setup)
        self.Vector_E = NamedCalcObject("Vector_E", setup)
        self.Vector_H = NamedCalcObject("Vector_H", setup)
        self.Vector_Jsurf = NamedCalcObject("Vector_Jsurf", setup)
        self.Vector_Jvol = NamedCalcObject("Vector_Jvol", setup)
        self.ComplexMag_E = NamedCalcObject("ComplexMag_E", setup)
        self.ComplexMag_H = NamedCalcObject("ComplexMag_H", setup)
        self.ComplexMag_Jsurf = NamedCalcObject("ComplexMag_Jsurf", setup)
        self.ComplexMag_Jvol = NamedCalcObject("ComplexMag_Jvol", setup)
        self.P_J = NamedCalcObject("P_J", setup)

        self.named_expression = {}  # dictionary to hold additional named expressions

    def clear_named_expressions(self):
        self.parent.parent._fields_calc.ClearAllNamedExpr()

    def declare_named_expression(self, name):
        """"
        If a named epression has been created in the fields calculator, this
        function can be called to initialize the name to work with the fields object
        """
        self.named_expression[name] = NamedCalcObject(name, self.setup)

    def load_named_expression(self, path):
        '''
        Loads a .clc file given by the designated path. Must be formatted correctly
        '''

        if check_path(path):
            load_file=open(path, 'r')
            lines=load_file.readlines()
            load_file.close()
            name=None
            for line in lines:
                if 'Name(' in line:
                    name=line.split('(\'')[1].split('\')')[0]
            if name==None:
                warnings.warn('Unable to find expression name in .clc file')
                return None
            else:
                try:
                    self.parent.parent._fields_calc.LoadNamedExpressions(path, 'Fields', name)
                    self.named_expression[name]=NamedCalcObject(name, self.setup)
                    return name
                except:
                    warnings.warn('Unable to load .clc file')
                    return None
        else:
            warnings.warn('.clc file requested not in path')
            return None

    def use_named_expression(self, name):
        """
        Expression can be used to access dictionary of named expressions,
        Alternately user can access dictionary directly via named_expression()
        """
        return self.named_expression[name]


class CalcObject(COMWrapper):
    def __init__(self, stack, setup):
        """
        :type stack: [(str, str)]
        :type setup: HfssSetup
        """
        super(CalcObject, self).__init__()
        self.stack = stack
        self.setup = setup
        self.calc_module = setup.parent._fields_calc

    def _bin_op(self, other, op):
        if isinstance(other, (int, float)):
            other = ConstantCalcObject(other, self.setup)

        stack = self.stack + other.stack
        stack.append(("CalcOp", op))
        return CalcObject(stack, self.setup)

    def _unary_op(self, op):
        stack = self.stack[:]
        stack.append(("CalcOp", op))
        return CalcObject(stack, self.setup)

    def __add__(self, other):
        return self._bin_op(other, "+")

    def __radd__(self, other):
        return self + other

    def __sub__(self, other):
        return self._bin_op(other, "-")

    def __rsub__(self, other):
        return (-self) + other

    def __mul__(self, other):
        return self._bin_op(other, "*")

    def __rmul__(self, other):
        return self*other

    def __div__(self, other):
        return self._bin_op(other, "/")

    def __rdiv__(self, other):
        other = ConstantCalcObject(other, self.setup)
        return other/self

    def __pow__(self, other):
        return self._bin_op(other, "Pow")

    def dot(self, other):
        return self._bin_op(other, "Dot")

    def __neg__(self):
        return self._unary_op("Neg")

    def __abs__(self):
        return self._unary_op("Abs")

    def __mag__(self):
        return self._unary_op("Mag")

    def mag(self):
        return self._unary_op("Mag")

    def smooth(self):
        return self._unary_op("Smooth")

    def conj(self):
        return self._unary_op("Conj")  # make this right

    def scalar_x(self):
        return self._unary_op("ScalarX")

    def scalar_y(self):
        return self._unary_op("ScalarY")

    def scalar_z(self):
        return self._unary_op("ScalarZ")

    def norm_2(self):

        return (self.__mag__()).__pow__(2)
        # return self._unary_op("ScalarX")**2+self._unary_op("ScalarY")**2+self._unary_op("ScalarZ")**2

    def real(self):
        return self._unary_op("Real")

    def imag(self):
        return self._unary_op("Imag")

    def complexmag(self):
        return self._unary_op("CmplxMag")

    def _integrate(self, name, type):
        stack = self.stack + [(type, name), ("CalcOp", "Integrate")]
        return CalcObject(stack, self.setup)

    def _maximum(self, name, type):
        stack = self.stack + [(type, name), ("CalcOp", "Maximum")]
        return CalcObject(stack, self.setup)

    def getQty(self, name):
        stack = self.stack + [("EnterQty", name)]
        return CalcObject(stack, self.setup)

    def integrate_line(self, name):
        return self._integrate(name, "EnterLine")

    def integrate_line_tangent(self, name):
        ''' integrate line tangent to vector expression \n
            name = of line to integrate over '''
        self.stack = self.stack + [("EnterLine", name),
                                   ("CalcOp",    "Tangent"),
                                   ("CalcOp",    "Dot")]
        return self.integrate_line(name)

    def line_tangent_coor(self, name, coordinate):
        ''' integrate line tangent to vector expression \n
            name = of line to integrate over '''
        if coordinate not in ['X', 'Y', 'Z']:
            raise ValueError
        self.stack = self.stack + [("EnterLine", name),
                                   ("CalcOp",    "Tangent"),
                                   ("CalcOp",    "Scalar"+coordinate)]
        return self.integrate_line(name)

    def integrate_surf(self, name="AllObjects"):
        return self._integrate(name, "EnterSurf")

    def integrate_vol(self, name="AllObjects"):
        return self._integrate(name, "EnterVol")

    def maximum_vol(self, name='AllObjects'):
        return self._maximum(name, 'EnterVol')

    def times_eps(self):
        stack = self.stack + [("ClcMaterial", ("Permittivity (epsi)", "mult"))]
        return CalcObject(stack, self.setup)

    def times_mu(self):
        stack = self.stack + [("ClcMaterial", ("Permeability (mu)", "mult"))]
        return CalcObject(stack, self.setup)

    def write_stack(self):
        for fn, arg in self.stack:
            if np.size(arg) > 1 and fn not in ['EnterVector']:
                getattr(self.calc_module, fn)(*arg)
            else:
                getattr(self.calc_module, fn)(arg)

    def save_as(self, name):
        """if the object already exists, try clearing your
        named expressions first with fields.clear_named_expressions"""
        self.write_stack()
        self.calc_module.AddNamedExpr(name)
        return NamedCalcObject(name, self.setup)

    def evaluate(self, phase=0, lv=None, print_debug=False):  # , n_mode=1):
        self.write_stack()
        if print_debug:
            print('---------------------')
            print('writing to stack: OK')
            print('-----------------')
        #self.calc_module.set_mode(n_mode, 0)
        setup_name = self.setup.solution_name

        if lv is not None:
            args = lv
        else:
            args = []

        args.append("Phase:=")
        args.append(str(int(phase)) + "deg")

        if isinstance(self.setup, HfssDMSetup):
            args.extend(["Freq:=", self.setup.solution_freq])

        self.calc_module.ClcEval(setup_name, args)
        return float(self.calc_module.GetTopEntryValue(setup_name, args)[0])


class NamedCalcObject(CalcObject):
    def __init__(self, name, setup):
        self.name = name
        stack = [("CopyNamedExprToStack", name)]
        super(NamedCalcObject, self).__init__(stack, setup)


class ConstantCalcObject(CalcObject):
    def __init__(self, num, setup):
        stack = [("EnterScalar", num)]
        super(ConstantCalcObject, self).__init__(stack, setup)


class ConstantVecCalcObject(CalcObject):
    def __init__(self, vec, setup):
        stack = [("EnterVector", vec)]
        super(ConstantVecCalcObject, self).__init__(stack, setup)


def get_active_project():
    ''' If you see the error:
        "The requested operation requires elevation."
        then you need to run your python as an admin.
    '''
    import ctypes
    import os
    try:
        is_admin = os.getuid() == 0
    except AttributeError:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
    if not is_admin:
        print('\033[93m WARNING: you are not runnning as an admin! \
            You need to run as an admin. You will probably get an error next.\
                 \033[0m')

    app = HfssApp()
    desktop = app.get_app_desktop()
    return desktop.get_active_project()


def get_active_design():
    project = get_active_project()
    return project.get_active_design()


def get_report_arrays(name: str):
    d = get_active_design()
    r = HfssReport(d, name)
    return r.get_arrays()


def load_ansys_project(proj_name: str, project_path: str = None, extension: str = '.aedt'):
    '''
    Utility function to load an Ansys project.

    Args:
        proj_name : None  --> get active. (make sure 2 run as admin)
        extension : `aedt` is for 2016 version and newer
    '''
    if project_path:
        # convert slashes correctly for system
        project_path = Path(project_path)

        # Checks
        assert project_path.is_dir(), "ERROR! project_path is not a valid directory \N{loudly crying face}.\
            Check the path, and especially \\ charecters."

        project_path /= project_path / Path(proj_name + extension)

        if (project_path).is_file():
            logger.info('\tFile path to HFSS project found.')
        else:
            raise Exception(
                "ERROR! Valid directory, but invalid project filename. \N{loudly crying face} Not found!\
                     Please check your filename.\n%s\n" % project_path)

        if (project_path/'.lock').is_file():
            logger.warning(
                '\t\tFile is locked. \N{fearful face} If connection fails, delete the .lock file.')

    app = HfssApp()
    logger.info("\tOpened Ansys App")

    desktop = app.get_app_desktop()
    logger.info(f"\tOpened Ansys Desktop v{desktop.get_version()}")
    #logger.debug(f"\tOpen projects: {desktop.get_project_names()}")

    if proj_name is not None:
        if proj_name in desktop.get_project_names():
            desktop.set_active_project(proj_name)
            project = desktop.get_active_project()
        else:
            project = desktop.open_project(str(project_path))
    else:
        project = desktop.get_active_project()
    logger.info(
        f"\tOpened Ansys Project\n\tFolder:    {project.get_path()}\n\tProject:   {project.name}")

    return app, desktop, project
