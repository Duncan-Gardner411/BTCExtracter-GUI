# BTCExtracter-GUI
This is a repository for sharing the program used to format BTC data in the standardized UNM format.
Contact duncangardner411@gmail.com for questions.

This includes two programs 
  BTC Extractor is a script to take BTC and context data and convert to UNM format
  BTC Correction is a script to correct any repeated mistakes by accepting folders containing the output of the BTC extractor and a set of fixing rules
  
These programs are written in python and require the following libraries.
    openpyxl
    tkinter
    
This also contains a database format file, which the BTC extractor uses as a baseline to add information to. Any changes to this format should also be replicated in the BTC extractor program.
    
These allow for the creation and manipulation of Excel sheets, and the creation of a UI
