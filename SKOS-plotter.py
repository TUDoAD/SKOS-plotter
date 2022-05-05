# -*- coding: utf-8 -*-
"""
Created on Fri Dec 31 15:41:14 2021

@author: Alexander Behr
@affiliation: TU Dortmund Fac. BCI - AG AD

Description:
This code first searches for all excel-files contained in subdirectory 
./import and then repeats the following loop for each file:
First, the URI_generation method from URIgenerator.py is applied 
(see description of that code for more details). 
Secondly, vocexcel module convert is applied, to convert the now conform excel-file
to a ttl-file, storing it in the ./export subdirectory.
Then the plots and documentations (both html-files) are generated and stored 
into ./export/dendro_<Excel-file-name> and ./export/docs_<Excel-file-name> respectively.
Those plots are done using Ontospy module.    
"""

# pip install ontospy
# pip install ontospy[FULL]

## full module (use this):
# pip install ontospy[FULL] -U
import ontospy
from ontospy.ontodocs.viz.viz_html_single import *
from ontospy.ontodocs.viz.viz_d3dendogram import *

import vocexcel
from vocexcel import convert

import glob
import os
from pathlib import Path
#import subprocess

import URIgenerator


filenames = glob.glob("./import/*.xlsx")


for name in filenames:
    print("name")
    URIgenerator.URI_generation(name)
    print("URIs generated and saved in {}".format(os.path.basename(name)))
    
    excel_path_with_URIs = "./export/" + os.path.basename(name)
    convert.excel_to_rdf(excel_path_with_URIs,output_file_path = "./export/")#,output_format = "xml")
    print("converted import/{} successfully to export/{}.".format(os.path.basename(name),Path(os.path.basename(name)).with_suffix(".ttl")))


    ##
    #       SKOS plotting begins
    ##
    model = ontospy.Ontospy("./export/"+str(Path(os.path.basename(name)).with_suffix(".ttl")), verbose = True)
    
    #model = ontospy.Ontospy("ExampleCFDOntology.owl", verbose = True)
    docs_model = HTMLVisualizer(model, title="docs")
    dendro = Dataviz(model, title = "dendrogram")
    
    docs_model.build(output_path='./export/docs_'+ os.path.splitext(Path(os.path.basename(name)))[0] + '/')
    dendro.build(output_path = './export/dendro_'+ os.path.splitext(Path(os.path.basename(name)))[0] + '/')

#dendro.preview()


#docs_model.preview()

#dendro.build()
#dendro.preview()


#ontospy.ontodocs.viz.viz_d3dendogram(model)