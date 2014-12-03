nlconverter
===========
A tool for converting e-mails from Lotus Notes .nsf  into .mbox format

Heavily based on https://code.google.com/p/nlconverter/

Main modifications are:
* modified regex in order to suit my needs
* uses extra Lotus Notes email fields (Inet*) in order to do better name/e-mail coupling. 
* does date parsing conformant to rfc5322
* enable different header charsets
* removes the .ical generation completely (not needed in my case)
* is command line only, does not contain the gui from the original project
