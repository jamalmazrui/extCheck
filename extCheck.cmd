@echo off
rem extCheck.cmd -- run extCheck from the command line, from the top of the
rem installed folder or of the project. The program lives in exec\, as in
rem every Homer app; arguments pass straight through.
"%~dp0exec\extCheck.exe" %*
