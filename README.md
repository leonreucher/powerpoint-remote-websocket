# powerpoint-remote-websocket

This is a VSTO PowerPoint plugin to remote control the slideshow features using a WebSocket connection.

This plugin is intended to get used by Bitfocus Companion. After installing this VSTO PowerPoint AddIn, Companion can communicate with PowerPoint to send commands and receive information.
But of course you can also use this plugin for other purposes.

### Installation

Go to the <a href="https://github.com/leonreucher/powerpoint-remote-websocket/releases" target="_blank">Releases</a> tab and download the latest version. 
After download the zip archive, make sure to right click the object and go to the file properties. In this window at the bottom check "Unblock" (otherwise you will get errors during the installation).

The last step is to unzip the downloaded file and to run the "installation.bat" script - now just follow the installer instructions.

### WebSocket API documentation

After you've setup the WebSocket connection, you can use the following command to retreive all information.
```json
{
  "action": "status"
}
```
The answer to this query will look like this:
```json
{
  "slideShowActive": boolean (R/W),
  "totalSlideCount": int (R),
  "currentSlide": int (R/W),
  "presentationFullPath": string (R),
  "fileName": string (R)
}
```

To start/stop the presentation or go to a specific slide, simply send the information in the following format:
```json
{
  "slideShowActive": true,
  "currentSlide": 4
}
```
The following actions are available:

Close all currently opened presentations:
```json
{
  "action": "closeAll"
}
```
Open a presentation from the filesystem:
```json
{
  "action": "openPresentation",
  "path": "C:\path\to\file\filename.pptx",
  "closeOthers": true
}
```
Go directly to first slide:
```json
{
  "action": "first"
}
```
Go directly to last slide:
```json
{
  "action": "last"
}
```
Go to next slide:
```json
{
  "action": "next"
}
```
Go to previous slide:
```json
{
  "action": "previous"
}
```
Blackout Presentation screen:
```json
{
  "action": "blackout"
}
```
Whiteout Presentation screen:
```json
{
  "action": "whiteout"
}
```
Show Presentation again after black-/whiteout:
```json
{
  "action": "showPresentation"
}
```
Hide slide:
```json
{
  "action": "hideSlide",
  "slideId": 1
}
```
Unhide slide:
```json
{
  "action": "unhideSlide",
  "slideId": 1
}
```
Unhide all slides:
```json
{
  "action": "unhideAllSlides"
}
```
Show laser pointer:
```json
{
  "action": "showLaserPointer"
}
```
Hide laser pointer:
```json
{
  "action": "hideLaserPointer"
}
```
Toggle laser pointer:
```json
{
  "action": "toggleLaserPointer"
}
```
Erase drawings on active slide:
```json
{
  "action": "eraseDrawings"
}
```
