import clr, os, winreg
from itertools import islice

# This boilerplate requires the 'pythonnet' module.
# The following instructions are for installing the 'pythonnet' module via pip:
#    1. Ensure you are running a Python version compatible with PythonNET. Check the article "ZOS-API using Python.NET" or
#    "Getting started with Python" in our knowledge base for more details.
#    2. Install 'pythonnet' from pip via a command prompt (type 'cmd' from the start menu or press Windows + R and type 'cmd' then enter)
#
#        python -m pip install pythonnet

# determine the Zemax working directory
aKey = winreg.OpenKey(winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER), r"Software\Zemax", 0, winreg.KEY_READ)
zemaxData = winreg.QueryValueEx(aKey, 'ZemaxRoot')
NetHelper = os.path.join(os.sep, zemaxData[0], r'ZOS-API\Libraries\ZOSAPI_NetHelper.dll')
winreg.CloseKey(aKey)

# add the NetHelper DLL for locating the OpticStudio install folder
clr.AddReference(NetHelper)
import ZOSAPI_NetHelper

pathToInstall = ''
# uncomment the following line to use a specific instance of the ZOS-API assemblies
#pathToInstall = r'C:\C:\Program Files\Zemax OpticStudio'

# connect to OpticStudio
success = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize(pathToInstall);

zemaxDir = ''
if success:
    zemaxDir = ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory();
    print('Found OpticStudio at:   %s' + zemaxDir);
else:
    raise Exception('Cannot find OpticStudio')

# load the ZOS-API assemblies
clr.AddReference(os.path.join(os.sep, zemaxDir, r'ZOSAPI.dll'))
clr.AddReference(os.path.join(os.sep, zemaxDir, r'ZOSAPI_Interfaces.dll'))
import ZOSAPI

TheConnection = ZOSAPI.ZOSAPI_Connection()
if TheConnection is None:
    raise Exception("Unable to intialize NET connection to ZOSAPI")

TheApplication = TheConnection.ConnectAsExtension(0)
if TheApplication is None:
    raise Exception("Unable to acquire ZOSAPI application")

if TheApplication.IsValidLicenseForAPI == False:
    raise Exception("License is not valid for ZOSAPI use.  Make sure you have enabled 'Programming > Interactive Extension' from the OpticStudio GUI.")

TheSystem = TheApplication.PrimarySystem
if TheSystem is None:
    raise Exception("Unable to acquire Primary system")

def reshape(data, x, y, transpose = False):
    """Converts a System.Double[,] to a 2D list for plotting or post processing
    
    Parameters
    ----------
    data      : System.Double[,] data directly from ZOS-API 
    x         : x width of new 2D list [use var.GetLength(0) for dimension]
    y         : y width of new 2D list [use var.GetLength(1) for dimension]
    transpose : transposes data; needed for some multi-dimensional line series data
    
    Returns
    -------
    res       : 2D list; can be directly used with Matplotlib or converted to
                a numpy array using numpy.asarray(res)
    """
    if type(data) is not list:
        data = list(data)
    var_lst = [y] * x;
    it = iter(data)
    res = [list(islice(it, i)) for i in var_lst]
    if transpose:
        return self.transpose(res);
    return res
    
def transpose(data):
    """Transposes a 2D list (Python3.x or greater).  
    
    Useful for converting mutli-dimensional line series (i.e. FFT PSF)
    
    Parameters
    ----------
    data      : Python native list (if using System.Data[,] object reshape first)    
    
    Returns
    -------
    res       : transposed 2D list
    """
    if type(data) is not list:
        data = list(data)
    return list(map(list, zip(*data)))

print('Connected to OpticStudio')

# The connection should now be ready to use.  For example:
print('Serial #: ', TheApplication.SerialCode)

# Insert Code Here
##################
import pandas as pd
import math

LDE = TheSystem.LDE
num_surfaces = LDE.NumberOfSurfaces
print(f"Found {num_surfaces} surfaces")

def safe_get_property(obj, prop_name, cast=None, default=None):
    """
    Safely attempt to get obj.<prop_name>. If the .NET call raises or returns None,
    return default. Optionally cast the value (e.g., float or str).
    """
    try:
        val = getattr(obj, prop_name)
    except Exception as e:
        # Property access raised (common for some special rows) -> return default
        # Uncomment next line if you want to debug which properties fail:
        # print(f"DEBUG: getattr failed for prop '{prop_name}': {e}")
        return default
    # val may be a .NET null -> treat as None
    if val is None:
        return default
    # try casting
    try:
        if cast is not None:
            return cast(val)
        return val
    except Exception:
        return default

# Predefine the keys/columns (keeps CSV consistent)
columns = [
    "Surface", "Comment", "Type", "Radius", "Thickness", "Material",
    "SemiDiameter", "Conic", "A4", "A6", "A8"
]

lens_data = []

for i in range(1, num_surfaces + 1):
    surf = None
    try:
        surf = LDE.GetSurfaceAt(i)
    except Exception as e:
        print(f"⚠️ Could not GetSurfaceAt({i}): {e}")
        # append a row of defaults so indexing remains consistent
        lens_data.append({
            "Surface": i,
            "Comment": "",
            "Type": "",
            "Radius": float("nan"),
            "Thickness": float("nan"),
            "Material": "",
            "SemiDiameter": float("nan"),
            "Conic": float("nan"),
            "A4": float("nan"),
            "A6": float("nan"),
            "A8": float("nan"),
        })
        continue

    # Build the row using safe_get_property for each attribute
    row = {
        "Surface": i,
        "Comment": safe_get_property(surf, "Comment", cast=str, default=""),
        "Type": safe_get_property(surf, "Type", cast=str, default=""),
        "Radius": safe_get_property(surf, "Radius", cast=float, default=float("nan")),
        "Thickness": safe_get_property(surf, "Thickness", cast=float, default=float("nan")),
        "Material": safe_get_property(surf, "Material", cast=str, default=""),
        "SemiDiameter": safe_get_property(surf, "SemiDiameter", cast=float, default=float("nan")),
        "Conic": safe_get_property(surf, "Conic", cast=float, default=float("nan")),
        # initialize asphere coeffs
        "A4": float("nan"),
        "A6": float("nan"),
        "A8": float("nan"),
        "A10": float("nan"),
        "A12": float("nan"),
    }

    # Try to read Asphere object and coefficients, each access wrapped
    try:
        asp = None
        try:
            asp = surf.Asphere
        except Exception as e:
            # Some surfaces may throw on .Asphere access; ignore and keep NaNs
            # Uncomment to debug: print(f"DEBUG: surf.Asphere access failed for surface {i}: {e}")
            asp = None

        if asp is not None:
            # Coefficients sometimes indexed starting at 0 or 1 depending on API - use try/except
            try:
                row["A4"] = safe_get_property(asp, "Coefficient", cast=lambda fn: fn(4), default=float("nan"))
            except Exception:
                try:
                    # fallback: asp.Coefficient might be a callable method accessed differently
                    row["A4"] = float(asp.Coefficient(4))
                except Exception:
                    row["A4"] = float("nan")

            try:
                row["A6"] = safe_get_property(asp, "Coefficient", cast=lambda fn: fn(6), default=float("nan"))
            except Exception:
                try:
                    row["A6"] = float(asp.Coefficient(6))
                except Exception:
                    row["A6"] = float("nan")

            try:
                row["A8"] = safe_get_property(asp, "Coefficient", cast=lambda fn: fn(8), default=float("nan"))
            except Exception:
                try:
                    row["A8"] = float(asp.Coefficient(8))
                except Exception:
                    row["A8"] = float("nan")
    except Exception as e:
        # Catch-all to prevent single-surface errors from stopping the whole run
        print(f"⚠️ Unexpected error while reading Asphere for surface {i}: {e}")

    lens_data.append(row)

# Convert to DataFrame and export
df = pd.DataFrame(lens_data, columns=columns)
csv_path = os.path.join(os.getcwd(), "lens_data.csv")
df.to_csv(csv_path, index=False)

print(f"✅ Exported lens data for {num_surfaces} surfaces to {csv_path}")
print(df.head(15))