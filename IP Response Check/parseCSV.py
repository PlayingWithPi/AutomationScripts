# Imports required by program.
import csv
import subprocess

# Global Variables #
# Nothing of dire importance.
spare_line = "\n\n==========================================================\n\n"

# Opening the CSV file to read from.
new_list = csv.reader(open("input.csv"), delimiter=",")

# Main Code #
# Attempts a try/except to see if any errors in the CSV file to begin with.
try:
    for i in new_list:
        ip_addr = i[0]
        # Responses are created using the stdout flag.
        resp = subprocess.run("ping -a -n 2 " + ip_addr,
                              stdout=subprocess.PIPE)
        # Changed from bytes to string.
        decode_resp = str(resp.stdout.decode("utf-8"))
        output = ""
        # Checks with possible messages that may pop up, cleaned them up &
        # made them easier to read / understand.
        if "unreachable" in decode_resp:
            output += ip_addr + ", destination cannot be reached!" + spare_line
        elif "could not find host" in decode_resp:
            # Do not like this.... Want to use default error across two lines.
            output += "Ping request could not find host " + ip_addr + \
                "\nPlease check the name and try again." + spare_line
        else:
            # If nothing fails, code runs from here. (some formatting done.)
            output += decode_resp.strip() + spare_line

# Basic Error Handling #
# IndexError - in case someone does not start from the first line.
except IndexError:
    output += "There seems to be an error in your file. :(\n"
    output += "Do not leave spaces before/between lines!\n"
    output += "Start from Cell: A1."

# All other exceptions that I have not yet thought of so far.
except Exception as e:
    output += str(e) + "\n\n"
    output += "Congrats! You found a bug that wasn't noticed!"
    output += "Why'd you need to break it :(\n\n"
    output += "Send steps taken & error code. To reproduce & squash."
    # Shameless Plug.
    output += "https://github.com/PlayingWithPi/AutomationScripts/issues"

# The Output File #
# This can be changed / customised to wherever you need your output file to go.
# current location is the root of parseCSV.py file.
with open("output.txt", "w+", newline="") as fd:
    fd.write(output.rstrip())
