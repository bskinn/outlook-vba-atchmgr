# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/)
and this project adheres to [Semantic Versioning](http://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.0.0] - 2017-07-24

Existing production version released as open source.

### Current Features

File detachment:

 * Free selection of target folder for file detachment (one folder per
   execution of the tool)
 * Multiple files can be detached to the single folder in one execution
 * Red text block inserted into messages indicating name of the file
   prior to detachment and local destination, including a live hyperlink
   to the detached file

File (re)attachment:

 * User-selectable (re)attachment of any local files linked in a message
 * For files originally detached by the tool, the red text block is removed
   when the relevant file is reattached, and the option is available to
   delete the detached version of the file
 * For files not originally detached by the tool, no changes are made to
   the text of the message, and the option is NOT AVAILABLE to delete the
   file from disk

