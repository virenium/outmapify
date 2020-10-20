# outmapify

## Usage

This PoC relies heavily on pypiwin32 ([https://pypi.org/project/pypiwin32/](https://pypi.org/project/pypiwin32/) ) package and credits to the maintainers for their work. 

To run this, PyInstaller ([https://www.pyinstaller.org/](https://www.pyinstaller.org/)) would be required to compile an executable for Microsoft Windows. 

![Usage](img/image1.png?raw=true)

This script will go through all the mailboxes a user can access and print the subject of the email to console and write body of the email to a file called output.txt in your current folder. 

## Consideration

Based on the outlook configuration, the following may pop up on the victim side. 

![Popup message](img/image2.png?raw=true)

This can be suppressed by modifying certain registry values. I will leave this to your creativity.

