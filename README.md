# Automated RF Design Demo-MASTER
These notebooks provide a (somewhat) complete compilation of the setup and implementation of the most commonly used utilities in Ansys HFSS for developing 2D and 3D cQED platforms. For 3D there are demos for:

* Eigenmode
* Q3D
* Driven Modal
* Eigenmode+EPR

There is also a demonstration of how to import popular file formats (.STP and .DXF for 3D and 2D respectively) and generating parts using the companion PyInventor API module (https://github.com/AndrewOriani/PyInventor). The Ansys functionality is derived from the Ansys module within pyEPR (https://github.com/zlatko-minev/pyEPR). These demos however will not work with the current stable release of pyEPR (0.8.4) and requires additional functionality that was added by me to successfully run the demonstrations. You can find this modified version inside of the pyEPR-UPDATED folder. To install this open a shell or cmd line at the file path and simply run:

'''
python setup.py install
''' 

This will automatically overwrite any existing pyEPR install with this modified version. NOTE: This is backwards compatible with pyEPR versions 0.8.2 or ealier if you have existing pyEPR code. 

This code also makes use of the 'ezdxf' and 'asteval' libraries, both of which can be installed via pip or conda. Otherwise the code uses the same dependancies as the master pyEPR module.

# Additional Ansys Functionality

In addition to the existing Ansys module within pyEPR, this modified version has added:
* Importation of design files
* Selection of faces or objects based on location
* The ability to return variable values in the specified user variables (currently only for length, general version to come)

 These additions allow for the use of the module as a more general design and setup utility on top of the existing EPR functionality.
