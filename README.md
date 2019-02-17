# BackgroundServer
A background server that allows you to manage your computer via a browser. Written in VB6.

# Functions
**Password Protected**

Password is required to manage your computer via a browser. You can set your own password.

**Process Manager**

You are able to view a full list of running processes with their information including PID, Parent PID, No. of threads, Priority and Image Path. You can select multiple processes to kill them, suspend them or resume them.

**File Manager**

You can view your files on a browser. The table shows file size, creation time, modification time, and date accessed. You can download files, too.

**Window List**

A list of all windows can be viewed. You can know which window is currently focused. The list includes whether the window is visible or not and Class Names, Window Captions and belonging process PIDs as well. You can do operations including Minimizing/Maximizing the window, Hiding/Showing the window, and closing the window. Also, you can view child windows of a parent window.

**Clipboard**

The program records the last change time of the system clipboard. You can view both of the text and image content in the clipboard. You can set the text content of clipboard, as well.

**VBS Scripting**

You can type a VBS script in the browser and run it on the remote computer.

**Command Line**

You can type a commandline in the browser and execute it on the remote computer. You can view the output of the command.

**Date & Time**

This page shows the date and time of the remote computer.

**Idle Time**

You can know the System Up Time and the last time that the user uses keyboard, mouse and changes window focus.

**Block Input**

You are able to lock/unlock the remote keyboard and mouse.

**Send Keys**

You can send keyboard input to the remote computer. You can choose either using keybd_event() function or the conventional SendKeys().

**Mouse Control**

You can send mouse input such as wheel up/down, left button up/down and so on. You can set the position of the cursor. What's more, there's a funny function which allows you to move the remote cursor crazily.

**Capture Screen**

You can view the screen of the remote computer in your browser.

**Settings**

The settings page shows the start time of the server program and you can modify the screen capture quality and maximum connection count. You click the "Log" button to show server connection logs and deauthenticate other users. You may restart the server by clicking "Restart Server" on the page. You can hide/show the server window and tray icon.

**MISC**

There's a "Blue Screen" button which makes BSOD on the remote computer.

The server records all connected IPs and shows network traffic. 

# Known Issues

You can't download files or view remote screen because I didn't send proper HTTP headers (#^.^#)

Sometimes blue screen doesn't work properly, IDK why XD

I don't know, maybe you can tell me... but I won't fix them cuz this project was finished last year and I don't want to touch those code anymore! XDDDDD

# License
MIT
