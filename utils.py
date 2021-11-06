import sys, os, re

def resource_path(relative_path):
    """
    Get absolute path to resource, works for dev and for PyInstaller

    Parameters
    ----------
    relative_path : str
        Relative path string to convert to absolute
    
    Returns
    -------
    str :
        Absolute path from `relative_path`
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def hex_to_RGB(color):
    """
    Converts a hex `color` to an RGB tuple

    Parameters
    ----------
    color : str
        Color hex string in the form "#000000"
    
    Returns
    -------
    tuple :
        Length-3 tuple of integers representing the RGB values for the hex `color`
    """
    no_hash = color.lstrip('#')   #Remove the '#' from the hex value
    return tuple(int(no_hash[i:i+2], 16) for i in (0, 2 ,4))

def use_dark(RGB_color):
    """
    Contrast checking formula from https://stackoverflow.com/questions/3942878/how-to-decide-font-color-in-white-or-black-depending-on-background-color
    
    Parameters
    ----------
    RGB_color : tuple
        Length-3 tuple of integer/float RGB values
    
    Returns
    -------
    bool :
        Whether black should be used against this color background
    """
    return (RGB_color[0]*0.299 + RGB_color[1]*0.587 + RGB_color[2]*0.114) > 186

def check_hex_color(color):
    """
    Checks if a string is a hex color

    Parameters
    ----------
    color : str
        String to check if it's a hex color
    
    Returns
    -------
    bool :
        Whether the string is a hex color
    """
    return bool(re.search(r'^#(?:[0-9a-fA-F]{3}){1,2}$', color))

def convert_to_type(obj, default, type_=int):
    """
    Trys to convert `obj` to `type`, and uses `default` if that fails

    Parameters
    ----------
    obj : object
        Object to convert
    default : object
        Default value if conversion fails
    type_ : class
        Class to convert the `obj`ect to
    """
    assert isinstance(default, type_), 'The default value must be of the same type or a subclass of the type'
    try:
        return type_(obj)
    except (ValueError, TypeError):
        return default
