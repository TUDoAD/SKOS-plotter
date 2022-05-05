This repository contains the SKOS-plotter by Alexander S. Behr, TU Dortmund University, created in context of NFDI4Cat TA1.

** Needed modules (list may not be complete!):**
	- pip install ontospy[FULL] -U

Please make sure to include at least one xlsx file in the import subdirectory. 
All files placed here should be conform with the SKOS-template we're using in NFDI4Cat.

## SKOS-plotter.py (Main program)
This code first searches for all excel-files contained in subdirectory ./import and then repeats the following loop for each file:
- First, the URI_generation method from URIgenerator.py is applied 
(see description of that code for more details). 
- Secondly, vocexcel module convert is applied, to convert the now conform excel-file
to a ttl-file, storing it in the ./export subdirectory.
- Then the plots and documentations (both html-files) are generated and stored 
into ./export/dendro_<Excel-file-name> and ./export/docs_<Excel-file-name> respectively.
Those plots are done using Ontospy module.

## URIgenerator.py
First method URI_generation reads in the URI for the vocabulary (in cell B1) 
and the preferred labels. To get a URI-valid name, the space characters are 
then replaced by an underscore and put concatenated with the vocabulary URI 
into the fields of concept URI. While this for loop runs, all preferred labels 
are stored with their respective URIs into a dictionary (URI_dict).
The second loop then iterates through columns 5 and 6 (children and related).
When an entry is found, that is not None and does not contain some of 
(https://, http:// or www.), the entry is split by commas and each of those 
subentries is then replaced by it's URI by the dictionary (URI_dict).


## ./export/
Please make sure to extract zip files, before trying to open dendrograms created by this tool.