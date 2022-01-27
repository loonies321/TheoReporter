# TheoReporter

## About
This parser is designed to create reports in xls format from reports from the equipment of a chemical synthesis laboratory during the production of radiopharmaceuticals.
## The parser uses pdf reports from the following pieces of equipment:
### synthesis module
* GE FastLab2 
### automatic drug filling system
* Comecer Clio
* Comecer Theodorico 2

## data that is being parsed
* activity of isotopes from cyclotron
* drug lot number
* date of manafacture
* type of drug (FDG/PSMA/FES/FET/FLT/DOPA)
* serial in format -nubmer-type-date NNNTTTTDDMMYY or NNNTTTDDMMYY
* duration of sinthesis
* activity of bulc
* volume of bulc
* drug certification time
* number of passport
* name of vial
* activity of vial
* volume of vial
* drug user name
* drug order date
* issuing time
* activity of remains (theodorico report only)
* volume of remains (theodorico report only)


## making exe file of parser
* download repository
* install requirements
* open cmd in your repository folder
* #pyinstaller myscript.py -F -n NAME_OF_EXE
for more information of pyinstaller use https://pyinstaller.readthedocs.io/en/stable/usage.html

## test files
i can give you reports pdfs files under NDA
# LICENSE
Copyright [2021] [Maksim Trushin]

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
