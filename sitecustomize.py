"""
This sitecustomize.py must be in the path in order for the current interpreter to connect with the
python process deployed by the current instance of RayStation. Any python process will then run this
before any other imports.

It is only required for debugging and should not be included in a package release.

This code has been modified from the module connect.connect_cpython.py which fails to test all possible
session keys due to the try/except statement at line 798. The SystemError exception does not catch the
actual exception and so doesn't increment the counter. The max counter value (10) is also too low.
"""

import os
import re
import clr

clr.AddReference('ScriptClient')
import ScriptClient


def test_pid_(pid_):
    """Connect function for connecting to a scripting service host."""
    # print(f"Attempting to connect to RayStation session id = {pid_}")
    try:
        test_instance = ScriptClient.RayScriptService.Connect(
            f'net.pipe://localhost/raystation_{pid_}'
        ).Instance
        test_instance.Client.GetCurrent('MachineDB')
        print(f'Script successfully connected to pid {pid_}.')
        return pid_
    except Exception as e:
        # print(f'Script failed to connect to RayStation. Error: {e}')
        return None


def set_raystation_pid():
    """Gets and sets the pid of a live RayStation instance."""
    raystation_pid = None
    for line in os.popen("tasklist").readlines():
        pttrn = r'.*RayStation.exe\s+(\d+)\s+.*'
        match = re.match(pttrn, line)
        if match:
            raystation_pid = match.group(1)
            break

    if raystation_pid is not None:
        for n in range(1, 40):
            pid = f"{raystation_pid}_{n}"
            print(f'Testing pid {pid}')
            if test_pid_(pid) is not None:
                os.environ['RAYSTATION_PID'] = pid
                print(f'Set RAYSTATION_PID to {pid}')
                return True

    print('Could not connect to RayStation.')
    return False


# Run the connection setup
set_raystation_pid()
