# 6.835-project

## What's in here:
packages/leap : folder of all the files needed to interface the LeapMotion sensor with python
help.pdf: help document for user actions/commands
requirements.txt: python 2.7 packages to install
main.py: the script for running the system


## Prereqs for running: 
Windows, Leapmotion 3.2.1 SDK (https://developer.leapmotion.com/releases/leap-motion-orion-321), Python 2.7 64 bit

## To run:
Copy the contents of packages/leap into the main directory where main.py is

Install the requirements in requirements.txt via pip2 (pip2 install -r requirements.txt)

Connect the leapmotion controller and place with the light facing towards you

Run main.py (python2 main.py)

## Tips
Sometimes the controller is finnicky, so it helps to have the diagnostic visualizer in the corner to see if the hands are being properly recognized

The controller might also refuse to connect to the app, it helps to run as an administrator or to close and reopen the terminal
