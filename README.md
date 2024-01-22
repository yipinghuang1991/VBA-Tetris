I am writing a Tetris game using Excel and VBA.
So that you have something to do if the IT policy in your company prevent you do install games.
One distinction of my attempt to others' is that almost everything is calculated when needed.
Feedbacks are welcome!

# Screenshot
![Screenshot](https://raw.githubusercontent.com/yipinghuang1991/VBA-Tetris/main/src/screenshot.png)

# Requirement
* Windows 10 64bit (Not tested on other platform)
* 64bit Microsoft Excel from 2016 up (32bit not tested)
* (Not required) Do not have other Excel running at the same time

# How To Play
1. Choose to enable macros when opening the file
2. Press START buttom to start the game
3. Press (and hold) Left arrow key to move left
4. Press (and hold) Right arrow key to move right
5. Press (and hold) Down arrow key to move downwards
6. Press Up arrow or X to rotate clockwise
7. Press Control or Z to rotate counterclockwise
7. Press Shift or C to hold piece
8. Press ESC to end game

# What's Working
* Vanilla Level 1 Tetris game (1000 millisecond gravity)
* Hold piece
* Pause/Resume
* View 6 incoming shapes
* 200 millisecond repeat delay
* 35 millisecond repeat rate

# To Do
* Implement score and level
* Add setting panel
* Add custom keybinding
* Add hard drop
* Add kick
* Add ghost piece
* Gerenal performance improvement
* Try to follow the [Tetris Guideline](https://tetris.fandom.com/wiki/Tetris_Guideline)
