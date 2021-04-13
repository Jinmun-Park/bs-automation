# Project Overview


## Project Background
<p align="justify"> Closing Assessment is a process of consolidating manual inputs from various levels of organization (Market/ Geo/ WW) using same excel format templates. Data submission from market/geo will be consolidated and and reviewed at WW level. It has been spending Numerous hours to collect manual inputs and place them into right format during critical closing period. Manual approach in consolidating the inputs and copying them across templates is very slow and it often results to have minor error due to manual copy-paste effort </p>

***

## Problem Statement
<p align="justify"> Current templates are standard on WW level, however Geo's using different ways to run consolidation of their markets. For example EMEA, AP & GCG, NA are using different templates for their markets and for more than 10 years, with small changes. Old templates are used, that can be obsolete, including a  lot of copy paste effort, often not reflecting requirements from GEO/ WW. The templates can also be adjusted each closing, causing further issues with month to month comparisons. </p>

***

## Project Goal
<p align="justify"> Redesign and update standard WW template to allow drill down to Market and Geo level and consolidation them automatic way. A process of consolidating each market/Geo templates will be run by execution file using Python. The automation will deploy the standard solution to all Geos and Markets. </p>

***

# Python Script
<p align="justify">Python scripts uses xlwings package to open excel template and perform copy/paste process. Yaml file is connected to main script allows users to change both directory path and range of cell for copy/paste. Log in and Log out function enables only authorized user to consolidate the selected markets. </p>

* **Configuration.yaml** : Any changes in directory path need to be typed in this file. This enable users to run executive file anywhere in desktop.

* **Setting.yaml** : The range of cells in excel template can key in here. If any excel sheets that needs to increase the size of rows and column can be adjusted in this yaml file.

* **Log** : Any errors while running executive file will be written in this file. 

<p align="justify"> auto-py-to-exe package from PyPi is converting package used in this project. Blank anaconda environment is made prior to execute conversion as it helps end users to perform consolidation quickly. Default environment may pack with your pre-installed packages that running executive file takes longer. </p>

<p align="justify"> All excel sheet names should be stored in dictionary. Dictionary enables to copy and paste through loop function in each excel worksheet sequentially. This is a list of excel sheets need to be stored in dictionary.</p>

* **Main** : Excel sheet where actuals from WWConsol and WSB are summarized in this tab. All manual adjustment values are stored and consolidated. 

* **Bridge** : Bridge tab where manual bridging inputs are saved here.
* **Roadmap** : Manual roadmap inputs are stored here.
* **Signing** : Manual signing inputs are stored here.
* **To-Go** : Manual to-go inputs are stored here.

* **Revenue** : This excel sheet is only existing in both Geo/WW level templates where all market revenue values are stored. Only WW consolidator can pull values from Geo template and store market information into 'WW_Rev' excel sheet.

