#--------------------------------#
# Name:         outmapify.py
# Purpose:      Search emails in all folders for keywords#
# Author:       vireniumk #
# 
# Usage:        outmapify.exe pattern1 pattern2 ... patternN
#
# This script uses MAPI to interface with the Outlook process and 
# searches user supplied key phrases in the emails. Outputs message
# body to "output.txt".
#
#--------------------------------#

from win32com.client import constants
from win32com.client.gencache import EnsureDispatch as Dispatch
import sys
import logging

logging.basicConfig(filename='log.txt', level=logging.DEBUG, 
                    format='%(asctime)s %(levelname)s %(name)s %(message)s')
logger=logging.getLogger(__name__)

outlook = Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")

class OutlookReader():
    def __init__(self, outlook_object):
        self._obj = outlook_object

    def items(self):
        array_size = self._obj.Count
        for item_index in range(1,array_size+1):
            yield (item_index, self._obj[item_index])

    def prop(self):
        return sorted( self._obj._prop_map_get_.keys() )


def printbanner():
    print("-"*100)
    print(" Name:         outmapify.py ")
    print(" Purpose:      Search emails in all folders for keywords ")
    print(" Author:       vireniumk ")
    print("")
    print("")
    print(" Usage:        outmapify.exe pattern1 pattern2 ... patternN ")
    print("-"*100)
    print("")
    print("")

def main():
    printbanner()
    if len(sys.argv)<2:
        sys.exit()

    f = open("output.txt","w+")
    for inx, folder in OutlookReader(mapi.Folders).items():
        # iterate all Outlook folders
        print ("-"*100)
        print ("[+] ",folder.Name)
        try:
            for inx,subfolder in OutlookReader(folder.Folders).items():
                messages=subfolder.Items
                try:
                    for message in messages:
                        subject = message.Subject
                        body = message.Body
                        if any(x.lower() in body.lower() for x in sys.argv):
                            print("Found in ",subject)
                            f.write(body)
                            f.write("-"*100)
                except Exception as e:
                    logger.error(e)
                    continue
        except Exception as e:
            logger.error(e)
            continue
    f.close()

if __name__ == '__main__':
    main()