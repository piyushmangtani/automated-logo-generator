#!/usr/bin/env python
# coding: utf-8

# In[2]:


import importlib
import subprocess

def check_and_install_libraries(required_libraries):
    missing_libraries = []

    for library in required_libraries:
        try:
            importlib.import_module(library)
        except ImportError:
            missing_libraries.append(library)

    if missing_libraries:
        print("The following libraries are missing and will be installed:")
        for library in missing_libraries:
            print(f" - {library}")

        install_libraries(missing_libraries)
    else:
        print("All required libraries are already installed.")

def install_libraries(libraries):
    for library in libraries:
        subprocess.call(["pip", "install", library])

if __name__ == "__main__":
    # List of required libraries
    required_libraries = ["python-pptx", "selenium", "webdriver-manager", "jupyterlab", "notebook", "pandas", "openpyxl","Pillow","fuzzywuzzy"]

    # Check and install libraries if needed
    check_and_install_libraries(required_libraries)

    # Rest of your script goes here
    # ...


# In[ ]:




