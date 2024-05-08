# W-IDM_OmeroImporterExcelHelperPy
The OMERO Importer Excel Helper Python tool was developed at **UMass Chan Medical School** and in collaboration with the **[Canada BioImaging](https://www.canadabioimaging.org)** Open Science Project for use with the Canada BioImaging [National OMERO Image Data Resource](https://omero.med.ualberta.ca/index/).

Its purpose is to create a list of Image data files and associated Tag annotations to be imported into OMERO using the [OmeroImporterPy](https://github.com/WU-BIMAC/W-IDM_OmeroImporterPy) tool. It automatically explores a target directory tree containing image data files to extract the names of the files and the names of the nested directories containing the files to be used as tag annotations at the image level in OMERO. 
The OMERO Importer Python Excel Helper tool can be run as a CLI command or it can be embedded in an Excel macro and run from a customized metadata-collection Excel spreadsheet.
