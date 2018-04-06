from cx_Freeze import setup, Executable

base = None


executables = [Executable("statementcalculator.py", base=base)]

packages = ["idna"]
options = {
    'build_exe': {

        'packages':packages,
    },

}

setup(
    name = "MTD Shift Parser",
    options = options,
    version = "1.0",
    description = 'Parses XML exports of MTD logon/logoff reports',
    executables = executables
)
