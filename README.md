# Iniseg DTP and Storyline

These are macros and regEx patterns I developed for Iniseg to better work with Word and Articulate Storyline.

## Word

The macros in Iniseg.bas have dependencies on [RaMacros.bas](https://github.com/RaRodRos/word-macros/blob/master/RaMacros.bas)

## GrepWin

Bookmarks file with regex patterns for editing the Storyline folders using grepWin.

The file must be named `bookmarks` and placed in `$env:APPDATA\grepWin\`

All the patterns are crammed into the same file for convenience, because grepWin can't detect files other names except `bookmarks` and it would be too tedious to change its name everytime it's going to be used.

## AutoHotKey

This script automatically inserts audio transcriptions in each slide in Storyline. It's not possible to do it programmatically, so It virtually clicks the buttons necessary for each slide.
