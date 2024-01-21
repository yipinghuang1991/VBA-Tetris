Disclaimer: This is a prototype. You may need to use task manager to kill excel when the unexpected happens. DO NOT RUN THIS GAME WITH OTHER UNSAVED EXCEL FILE OPEN. YOU HAVE BEEN WARNED.

I am writing an Tetris using Excel and VBA because why not.
One distinction of my attempt to others' is that almost everything is calculated when needed.
Feedbacks are welcome!

# Screenshot
![Screenshot](https://raw.githubusercontent.com/yipinghuang1991/VBA-Tetris/main/src/screenshot.png)

# Requirement
* Windows 10 64bit (Not tested on other platform)
* 64bit Microsoft Excel from 2016 up (32bit not tested)
* (Not required) Do not have other Excel running at the same time

# How To Play
1. When file is open, choose to accept running macro
2. Press START buttom to start game
3. Press (and hold) left arrow key to move left
4. Press (and hold) right arrow key to move right
5. Press (and hold) down arrow key to move downwards
6. Press up arrow key to rotate clockwise
7. Press H to hold piece
8. Press ESC to end game

# What's Working
* Vanilla Level 1 Tetris game (1000 millisecond gravity)
* Hold piece
* View 6 incoming shapes

# What's Broken
* You can hold piece repetitively
* Excel can freeze by unknown reasons

# To Do
* Improve README.md
* End game screen (when new piece cannot spawn, nothing happen)
* Implement score and level
* Add comments
* Add setting panel
* Add custom keybinding
* Add pause
* Add hard drop
* Add kick
* Add ghost piece
* Try to follow the [Tetris Guideline](https://tetris.fandom.com/wiki/Tetris_Guideline)
* Gerenal performance improvement
