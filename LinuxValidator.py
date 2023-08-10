# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import subprocess
import socket

class TextColors:
    RED = '\033[91m'
    GREEN = '\033[92m'
    RESET = '\033[0m'

def is_software_installed(partial_name):
    try:
        # For apt (Debian/Ubuntu based systems)
        result = subprocess.run(['dpkg', '-l'], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if partial_name.replace("_", "-") in result.stdout.decode():
            return True
    except:
        pass

    try:
        # For yum (RHEL/CentOS based systems)
        result = subprocess.run(['yum', 'list', 'installed'], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if partial_name in result.stdout.decode():
            return True
    except:
        pass

    return False

hostName = socket.getfqdn()

print(f"Scanning host: {hostName}")
software = ["ds_agent", "cortex", "opsramp"]
for app in software:
    if is_software_installed(app):
        print(f"{TextColors.GREEN}Software related to '{app}' is installed!{TextColors.RESET}")
    else:
        print(f"{TextColors.RED}No software related to '{app}' found.{TextColors.RESET}")