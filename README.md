# PyAPLUS

[![Documentation Status](https://readthedocs.org/projects/pyaplus-docs/badge/?version=latest)](https://pyaplus-docs.readthedocs.io/en/latest/?badge=latest)

Abstraction layer over the COM ASPEN PLUS interface using Python. If you can though, and I cannot recommend this enough, use HYSYS. The state is absolute pre-alpha right now, and it is only able to make reading some properties from streams quicker and to access them without that much code.

## Installation

Install the latest version of this repository to your machine following one of the options below accordingly to your preferences:

- users with git:<br/>
<pre>git clone https://github.com/DanielVazVaz/PyAPLUS.git
cd PyAPLUS
pip install -e .
</pre>

- users without git:<br/>
Browser to https://github.com/DanielVazVaz/PyAPLUS, click on the `Code` button and select `Download ZIP`. Unzip the files from your Download folder to the desired one. Open a terminal inside the folder you just unzipped (make sure this is the folder containing the `setup.py` file). Run the following command in the terminal:
<pre>
pip install -e .
</pre>

## Closing the simulation

Aspen Plus does not behave as nicely as HYSYS. Therefore, there is a forced closure programmed as the default that kills the task in the windows task manager, and all tasks with a shared name, i.e., AspenPlus.exe. If this is too extreme for you, you can remove this behavior by passing an argument to the `Simulation.close()` method, in the form `Simulation.close(soft = True)`. This probably will leave an Aspen Plus task in the background. If someone finds a less drastic way to make sure that Aspen Plus closes when you tell it to, I'm all ears.

## win32 DLL problem

Right now, it looks like for Python 3.8. there are problems with the `win32api` package. This worked for me:

```
pip install pywin32==225
```

If this still does not work, make sure that you do not have other `pywin32` in your environment, e.g., some version installed with `conda`.

## Win32 shenanigans: module win32com.gen_py has no attribute 'CLSIDToPackageMap'

If you get this error, which can occur from one day to the following, without having changed anything, in order to solve it, you must go to your temp folder. Normally, the path looks like:

```
C:\Users\<username>\AppData\Local\Temp
```

In there, there should be a folder named `gen_py` somewhere. Sometimes, it is inside another folder, such as `2\gen_py`. Delete it. This should fix the problem.

I have no idea why this happens, or what this is. 
