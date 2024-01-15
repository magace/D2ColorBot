# D2ColorBot
 Overlay for D2Bot Manager.  Used to change colors/dim the overlay.  Finally, my wife won't give me dirty looks when she's trying to sleep!
 It should auto popup the overlay when the manager window is active.

**Note:  This could use some work making the overlay fade or something cleaner.  I am not a programmer just here to learn a few things and have some fun.**
 
**How to install:**
Install python
Pip install the following packages:
pip install pygetwindow
pip install pynput
pip install rich
pip install pywin32

**Config notes**
delay is in seconds it's how long the loop is between checking windows.  Lower could spike cpu but too high will not be responsive enough.
window_title should be the title of the manager or at least the start of it so you could use a keyword like just "D2BOT #"

**Keyboard Commands**
left arrow: lowers transparency
right arrow: raises transparency
shift + left or right arrow: Cycles colors.

Images:
![image](https://github.com/magace/D2ColorBot/assets/7795098/9fcb7e4e-e0d9-4c25-812d-e8205c476180)
![image](https://github.com/magace/D2ColorBot/assets/7795098/b6fa230e-4984-4391-83c0-38004ec2fea8)
![image](https://github.com/magace/D2ColorBot/assets/7795098/8d9ab8a9-ddf2-4e78-82c1-39a2a5e88fb8)
![image](https://github.com/magace/D2ColorBot/assets/7795098/ae8b1d69-b249-4ac4-b22e-5230f3043d6c)
