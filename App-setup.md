Setup instructions for building .app from python script

✅ Step-by-Step: .app Packaging for Your PyQt6 Script
1. ✅ Install Dependencies
Make sure you're in a virtual environment, and install everything:
python3 -m venv venv
source venv/bin/activate
pip install py2app PyQt6 pandas numpy


2. ✅ Create setup.py
Here's a tailored setup.py for your app:
from setuptools import setup

APP = ['your_script.py']  # Replace with your actual filename
DATA_FILES = []
OPTIONS = {
    'argv_emulation': True,
    'includes': ['PyQt6', 'pandas', 'numpy'],
    'packages': ['PyQt6', 'pandas', 'numpy'],
    'resources': [],  # Add icons or data files here if needed
    'plist': {
        'CFBundleName': 'YourAppName',
        'CFBundleDisplayName': 'YourAppName',
        'CFBundleIdentifier': 'com.yourname.yourapp',
        'CFBundleVersion': '0.1.0',
        'CFBundleShortVersionString': '0.1.0',
    },
}

setup(
    app=APP,
    name='YourAppName',
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)

Replace:
your_script.py → your actual filename.


YourAppName and bundle identifiers accordingly.



3. ✅ Build the .app
In the terminal:
python setup.py py2app

If everything is configured correctly, this will generate:
dist/
  └── YourAppName.app


4. ✅ Launch Your App
open dist/YourAppName.app


🛠 Troubleshooting Tips
If PyQt6-related resources (like icons or styles) don't appear properly, run the app from Terminal using:

 ./dist/YourAppName.app/Contents/MacOS/YourAppName
 and read the console for any missing file errors.


For matplotlib, Qt plugins, or similar packages, you may need to add:

 'frameworks': ['/path/to/Qt/plugins'],
 in the OPTIONS.



