# powerpoint-remote-websocket

This is a VSTO PowerPoint plugin to remote control the slideshow features using a WebSocket connection.

This plugin is intended to get used by Bitfocus Companion. After installing this VSTO PowerPoint AddIn, Companion can communicate with PowerPoint to send commands and receive information.
But of course you can also use this plugin for other purposes.

### Installation

Go to the <a href="https://github.com/leonreucher/powerpoint-remote-websocket/releases" target="_blank">Releases</a> tab and download the latest version. Unzip the downloaded file and run the "installation.bat" script.

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
