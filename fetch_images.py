import sys
import os
import pandas as pd
import xlwings as xw
import json
import re

# maximum file name length
MAX_NAME = 255

#with a given directory eg. dataset5\Timepoint_1\ZStep_1
#extract the timepoint and ZStep. Covers all possible folder structure outcomes (missing ZStep folder or Timepoint folder...)
def getTandZ(dirpath):
    timepoint = ""
    zStep = ""
    parts = dirpath.split("\\")
    if len(parts) == 1:
        timepoint = "TimePoint_1"
        zStep = "ZStep_1"
    elif len(parts) == 3:
        timepoint= parts[1]
        zStep = parts[2]
    elif len(parts) == 2:
        if parts[1].split("_")[0] == "ZStep":
            zStep = parts[1]
            timepoint = "TimePoint_1"
        elif parts[1].split("_")[0] == "TimePoint":
            timepoint = parts[1]
            zStep = "ZStep_1"
    #return just the number of the timepoint and zStep
    return timepoint.split("_")[1], zStep.split("_")[1]

#This function returns a json that holds all possible wells/wavelengths/sites. Used for determining incomplet or missing wells
def getAllWells(htd):
    allWells = {}
    # number of sites in htd file
    siteCount = list(range(1,htd['sites']+1))
    # number of waves in htd file
    waveCount = htd['wavelength']['number']

    for i in range(1,htd['TimePoints']+1):
        timepoint = "TimePoint_"+str(i)
        allWells[timepoint] = {}
        for j in range(1,htd['ZSteps']+1):
            zStep = "ZStep_"+str(j)
            allWells[timepoint][zStep] = {}

            #populate the complete list of wells
            for well in htd['wells']:
                allWells[timepoint][zStep][well] = {}
                for i in range(waveCount):
                    allWells[timepoint][zStep][well]["w"+str(i+1)] = {
                        "Name": htd['wavelength']['names'][i],
                        "Sites" : list(siteCount)
                    }

    return allWells

#This is used to remove all used wells from the list of total wells, 
#leaving us with only the incomplete wells
def subtractJson(all,used):
    for timepoint in used:
        for zStep in used[timepoint]:
            for well in used[timepoint][zStep]:
            #if the well is used, check if all wavelengths are there
                for wavelength in used[timepoint][zStep][well]:
                #if the wavelength is used, check if all sites are there
                    for site in used[timepoint][zStep][well][wavelength]['Sites']:
                        #remove site
                        if (site in used[timepoint][zStep][well][wavelength]['Sites']):
                            all[timepoint][zStep][well][wavelength]['Sites'].remove(site)

                        #if the wavelength is now empty, remove it
                        if not all[timepoint][zStep][well][wavelength]['Sites']:
                            del all[timepoint][zStep][well][wavelength]
                        
                        #if the well is now empty, remove it
                        if not all[timepoint][zStep][well]:
                            del all[timepoint][zStep][well]
                        
                        #if the zStep is now empty, remove it
                        if not all[timepoint][zStep]:
                            del all[timepoint][zStep]
                        
                        #if the timepoint is now empty, remove it
                        if not all[timepoint]:
                            del all[timepoint]
    return all

# checks if an image is in the correct format. Returns the image if true
def checkName(filename,htd):
    #if there are no sites, then the regex is
    if htd["sites"] == 1:
        regex = "(.+)_([A-Z]\d+)_(w\d+).TIF"
        return re.match(regex,filename)

    #if there are no waves, then the regex is
    if htd["wavelength"]["number"] == 1:
        regex = "(.+)_([A-Z]\d+)_(s\d+).TIF"
        re.match(regex,filename)
    
    if htd["sites"] == 1 and htd["wavelength"]["number"] == 1:
        regex = "(.+)_([A-Z]\d+).TIF"
        return re.match(regex,filename)

    #if there are both sites and waves, then the regex is
    regex = "(.+)_([A-Z]\d+)_(s\d+)_(w\d+).TIF"
    return re.match(regex,filename)
    




#matches every image in a folder to a regex and adds each part to the valid or refused JSON object
#htd: the htd dictionary that is being used
def getImages(directory,htd):

    validImages = {}
    refusedImages = {}

    for filename in os.listdir(directory):
        itemPath = os.path.join(directory, filename)
        # if its a directory, check if its a timepoint or zstep
        if os.path.isdir(itemPath):
            if filename.split("_")[0] == "TimePoint":
                timePoint = filename
                validImages[timePoint] = {}
                refusedImages[timePoint] = {}

                # we now search through the timepoint folder
                for timepointFileName in os.listdir(itemPath):
                    timepointFilePath = os.path.join(itemPath, timepointFileName)
                    if os.path.isdir(timepointFilePath):
                        #find the zstep folder
                        #however, if the ZStep is 0, then we don't add it
                        #TODO process Zprojections
                        if timepointFileName.split("_")[0] == 'ZStep' and int(timepointFileName.split("_")[1]) != 0:
                            ZStep = timepointFileName
                            validImages[timePoint][ZStep] = {}
                            refusedImages[timePoint][ZStep] = {}

                            # we now search through the zstep folder, it should only contain images
                            for zstepFileName in os.listdir(timepointFilePath):
                                validImages, refusedImages = readImage(directory,validImages, refusedImages,zstepFileName,timePoint,ZStep,htd)

                    #ZStep is 1 if there are no folders within a timepoint folder
                    else:

                        #if the ZStep_1 doesn't exist yet, create it
                        ZStep = "ZStep_1"
                        if ZStep not in validImages[timePoint]:
                            validImages[timePoint][ZStep] = {}
                            refusedImages[timePoint][ZStep] = {}
                        validImages, refusedImages = readImage(directory,validImages, refusedImages,timepointFileName,timePoint,ZStep,htd)


            #if there are no timepoint folders, then it must be a zstep folder
            elif filename.split("_")[0] == 'ZStep':
                timePoint = "TimePoint_1"
                ZStep = filename
                if timePoint not in validImages:
                    validImages[timePoint] = {}
                    refusedImages[timePoint] = {}
                if ZStep not in validImages[timePoint]:
                    validImages[timePoint][ZStep] = {}
                    refusedImages[timePoint][ZStep] = {}

                for zstepFileName in os.listdir(itemPath):
                    validImages, refusedImages = readImage(directory,validImages, refusedImages,zstepFileName,timePoint,ZStep,htd)  

        # if it is a file
        elif os.path.isfile(itemPath):
            timePoint = "TimePoint_1"
            ZStep = "ZStep_1"
            if timePoint not in validImages:
                validImages[timePoint] = {}
                refusedImages[timePoint] = {}
                validImages[timePoint][ZStep] = {}
                refusedImages[timePoint][ZStep] = {}
            
            validImages, refusedImages = readImage(directory,validImages, refusedImages,filename,timePoint,ZStep,htd)  
    
    #remove empty timepoints if any
    validImages = cleanEmptyEntries(validImages)
    refusedImages = cleanEmptyEntries(refusedImages)
        
    return validImages, refusedImages

def cleanEmptyEntries(refusedImages):
    # Collect timepoints and z-steps to delete
    timepointsToDelete = []

    for timePoint in list(refusedImages.keys()):
        zstepsToDelete = []

        # Collect empty ZSteps to delete within the current timepoint
        for ZStep in list(refusedImages[timePoint].keys()):
            if refusedImages[timePoint][ZStep] == {}:
                zstepsToDelete.append(ZStep)

        # Delete empty ZSteps after collecting them
        for ZStep in zstepsToDelete:
            del refusedImages[timePoint][ZStep]

        # If the entire timepoint is empty, mark it for deletion
        if refusedImages[timePoint] == {}:
            timepointsToDelete.append(timePoint)

    # Delete empty timepoints after collecting them
    for timePoint in timepointsToDelete:
        del refusedImages[timePoint]

    # Return the cleaned dictionary
    return refusedImages

# sends an image to the valid or refused json object
#(valid or refused json objects must already be existing)
#TODO: check for valid/invalid timepoint and zStep
def readImage(directory,validImages, refusedImages,filename,timePoint,ZStep,htd):
     # verify the name matches the regex
    match = checkName(filename,htd)
    if match:
        # if there is 1 wavelength and/or site, then set it to 1 because it will not be included in the filename
        if htd["sites"] == 1 and htd["wavelength"]["number"] == 1:
            plateName, well, = match.groups()
            site = "s1"
            wavelength = "w1"
        elif htd["sites"] == 1:
            plateName, well, wavelength = match.groups()
            site = "s1"
        elif htd["wavelength"]["number"] == 1:
            plateName, well, site = match.groups()
            wavelength = "w1"
        else:
            plateName, well, site, wavelength = match.groups()
        
        # add the well if it doesn't exist in the valid json object
        if well not in validImages[timePoint][ZStep] and well in htd['wells']:
            validImages[timePoint][ZStep][well] = {}
        #add to refused images
        elif well not in refusedImages[timePoint][ZStep] and well not in htd['wells']:
            refusedImages[timePoint][ZStep][well] = {}


        #add the wavelength to the appropriate object if it has not been added yet
        if (well in validImages[timePoint][ZStep]):
            # add the wavelength if it does not exist and is valid
            if (wavelength not in validImages[timePoint][ZStep][well]) and (int(wavelength[1:]) <= htd['wavelength']['number']):
                validImages[timePoint][ZStep][well][wavelength] = {
                    "Name" : htd['wavelength']['names'][int(wavelength[1:]) -1],
                    "Sites" : {}
                }
        else:
            if (wavelength not in refusedImages[timePoint][ZStep][well]):
                refusedImages[timePoint][ZStep][well][wavelength] = {
                    #There is no name that exists for refused wells
                    "Sites": {}
                }
        #finally, add the site to the appropriate object if it has not been added yet
        if (int(site[1:]) <= htd['sites']) and (well in validImages[timePoint][ZStep]):
                if (int(site[1:]) not in validImages[timePoint][ZStep][well][wavelength]["Sites"]):
                    validImages[timePoint][ZStep][well][wavelength]["Sites"][int(site[1:])] = {
                        "filename": filename,
                        #TODO find out new filename
                        "newFilename": filename,
                        "filePath" : directory,
                        "Platename" : plateName,
                        "wellId" : well,
                        "siteId" : site,
                        "waveLengthId" : wavelength
                    }

        else:
            if (int(site[1:]) not in refusedImages[timePoint][ZStep][well][wavelength]["Sites"]):
                refusedImages[timePoint][ZStep][well][wavelength]["Sites"][int(site[1:])] = {
                    "filename": filename,
                    #TODO find out new filename
                    "newFilename" : filename,
                    "filePath" : directory,
                    "Platename" : plateName,
                    "wellId" : well,
                    "siteId" : site,
                    "waveLengthId" : wavelength
                }
    return validImages, refusedImages  


#get all wells that are incomplete within a HTD object
# htd: the htd file dictionary
# validWells: the dictionary containing all valid wells
def getIncompleteWells(htd,validWells):
    allWells = getAllWells(htd)
    incompleteWells = subtractJson(allWells,validWells)
    return incompleteWells

#takes the valid image dictionary and returns a list containing every image name
def getValidImageNames(validImages):
    imageList = []
    for timepoint in validImages:
        for zstep in validImages[timepoint]:
            for well in validImages[timepoint][zstep]:
                for wavelength in validImages[timepoint][zstep][well]:
                    for site in validImages[timepoint][zstep][well][wavelength]["Sites"]:
                        imageList.append(validImages[timepoint][zstep][well][wavelength]["Sites"][site]["filename"])
    return imageList

def parseContents(file):
    data = {}
    for line in file:
        line = line.strip()
        if line == '"EndFile"':
            break
        key, *value = line.split(', ')
        key = key.strip('"')

        # Handling boolean values
        value = [v.strip() == 'TRUE' for v in value] if any(v in ('TRUE', 'FALSE') for v in value) else [v.strip('"') for v in value]

        # Flatten lists if there's only one element
        if len(value) == 1:
            value = value[0]
        data[key] = value
    return data

# Read file contents of HTD and convert to JSON
def HTD_to_JSON(fileLocation):
    file = open(fileLocation, "r")
    parsedData = parseContents(file)

    # Convert the dictionary to JSON format
    jsonData = json.dumps(parsedData,indent=4)
    return json.loads(jsonData)

#returns a list of wavelength names
def getWaveLengthData(jsonFile):
    size = jsonFile.get("NWavelengths")

    # store the wavelength names in a list
    wavelengthNames = []

    # get the names of each wavelength
    for i in range(int(size)):
        wavelengthNames.append(jsonFile.get("WaveName"+str(i+1)))
    return wavelengthNames

#returns a list of the well names in use ex: A01, A03...
def getWells(jsonFile):
    wells = []

    #the column is the number
    xWells = jsonFile.get("XWells")

    #the row is the letter
    yWells = jsonFile.get("YWells")
    
    for i in range(int(yWells)):
        for j in range(int(xWells)):
            #check each well if it is used or not
            check = jsonFile.get("WellsSelection"+str(i+1))[j]
            if check is True:
                #we start at index 1 instead of 0, that's why we use j+1
                #get the corresponding letter for our number
                letter = chr(i + ord('A'))

                num = ""
                #make sure every number used is at least 2 digits. A1 -> A01
                if j+1 < 10:
                    num = str(0) + str(j+1)
                else:
                    num = str(j+1)
                wells.append(letter+num)
    return wells

# Constructs a json file based on a given HTD file
# (This is the function we will be calling)
def constructHTDInfo(fileLocation):
    #writes HTD data to a json object for processing
    data = HTD_to_JSON(fileLocation)

    wellsList = getWells(data)
    waveList = getWaveLengthData(data)

    #create JSON object for output
    info = {}
    info['wavelength'] = {"number":len(waveList), "names":waveList}
    info['wells'] = wellsList
    #check if there are multiple sites
    if data.get("XSites"):
        info['sites'] = int(data.get("XSites")) * int(data.get("YSites"))
    else:
        info['sites'] = 1
    
    #set zSteps, timepoints
    info['ZSteps'] = int(data.get("ZSteps"))
    info['TimePoints'] = int(data.get("TimePoints"))
    return info

# If present, returns the important HTD data in a dictionary
# location: folder location that will be searched
def getHtdFile(location):
    for file in os.listdir(location):
        filename, extension = os.path.splitext(file)
        if extension == ".HTD":
            return constructHTDInfo(os.path.join(location, file))
    #if there is no htd file, return None
    return None
        
# TODO: replace \\ with os.path.sep
def truncate_name(dirs, max=MAX_NAME):
    """Construct new file name from the end to the front until it exceeds MAX_NAME.

     Args:
        dirs: list of directory names ["tag1", "tag2", ..., "tagN", file].
        max: integer maximum length for file name.
     Returns:
        String: file name shorter than MAX_NAME
    """
    new_filename = os.path.sep.join(dirs)
    length = len(new_filename)
    num_split = 1
    while length > max:
        # split off front num_split directories
        split = new_filename.split(os.path.sep, num_split)
        length = len(split[-1])
        num_split += 1
    new_filename = split[-1]
    return new_filename

def walk_files(root, extensions, isSPW):
    """Navigates each file and folder beginning at the root folder and checks for image files.

    Args:
        root: path of the root project folder.
        extensions: list of string image file extensions.
        isSPW: boolean if the images are in the SPW format or not
    Yields:
        Yields file name and file path of each image file found.
    Raises:
    """
    for dirpath, dirnames, files in os.walk(root.split("\\")[-1], topdown=True):
        for file in files:
            if isSPW:
                timepoint, zStep = getTandZ(dirpath)
            filename, extension = os.path.splitext(file)
            if extension in extensions:
                filepath = os.path.join(dirpath, file)
                # get accompanying json file
                json = None
                for f in os.listdir(dirpath):
                    name, ext = os.path.splitext(f)
                    if (ext == ".json") and (name.lower().count(file.lower()) > 0):
                        json = f
                if isSPW:
                    yield (file, filepath, json, timepoint, zStep)
                else:
                    yield (file, filepath, json)

#creates a dataframe for images using the project/dataset structure
def create_DataFrame_ProjDataset(root, extensions, columns=["File Name","New File Name","File Path","MMA File Path","Tags"]):
    """Create a DataFrame containing image file info for the root directory.

    DataFrame will contain file name, new file name constructed by replacing backslashes in filepath
    with underscores, file path, MMA file path for the accompanying json file path, and tags which are the root directories subdirectories seperated by hashtags.

     Args:
        root: path to a dataset directory continaing subdirectories and image files.
        extensions: list of string image file extensions.
        columns: column names used to create the DataFrame and CSV output.
     Returns:
        A DataFrame containing information written to the output csv file.
     Catches:
        ValueError: catches value error from walk_files() if encountering a path longer than MAX_NAME characters
    """
    df = pd.DataFrame(columns=columns)
    walk = walk_files(root, extensions, False)

    for file, filepath, json in walk:
        # seperate the directories in the path from the ending file
        dirpath = os.path.dirname(filepath)

        # remove the root part of the dirname
        dirpath = dirpath.replace(root, "", 1)

        # split dirpath into list of directories
        dirs = []
        while 1:
            head, tail = os.path.split(dirpath)
            dirs.insert(0, tail)
            dirpath = head
            if (head == "") or (head == os.path.sep): # head=="" is first level file, head==os.path.sep is second or greater level file
                break

        # seperate tags with #
        tags = "#".join(dirs)

        dirs.append(file)

        # join directories with "_"
        new_filename = "_".join(dirs)

        # truncate file name if too long
        if len(new_filename) > MAX_NAME:
            new_filename = truncate_name(dirs)

        # add row to the DataFrame
        row = pd.DataFrame([[file, new_filename, filepath, json, tags]], columns=columns)
        df = df._append(row, ignore_index=True)

    df.sort_values(by="File Name", inplace=True)
    return df

#creates a dataframe for images using the SPW structure
def create_DataFrame_SPW(root,valid, extensions,htd, columns=["Well_Name","IMAGE NAME","File Name","File Path","MMA FILE PATH","Site ID", "Wavelength ID", "Z Score", "TimePoint"]):
    """Create a DataFrame containing image file info for the root directory.

    DataFrame will contain file name, new file name constructed by replacing backslashes in filepath
    with underscores, file path, MMA file path for the accompanying json file path, and tags which are the root directories subdirectories seperated by hashtags.

     Args:
        root: path to a dataset directory continaing subdirectories and image files.
        valid: a list of valid image names. Only images in this list will be added to the excel file
        extensions: list of string image file extensions.
        htd: the htd file dictionary used
        columns: column names used to create the DataFrame and CSV output.
     Returns:
        A DataFrame containing information written to the output csv file.
     Catches:
        ValueError: catches value error from walk_files() if encountering a path longer than MAX_NAME characters
    """
    df = pd.DataFrame(columns=columns)
    walk = walk_files(root, extensions, True)
    for file, filepath, json, timepoint, zStep in walk:

        #check if the image is valid before adding it to the excel file
        if file in valid:
            #TODO use this function when testing on a machine with omero installed in python
            #match = checkName(file)   

            #TODO will not work if htd is null
            match = checkName(file, htd)
            # if there is 1 wavelength and/or site, then set it to 1 because it will not be included in the filename
            if htd["sites"] == 1 and htd["wavelength"]["number"] == 1:
                plateName, well, = match.groups()
                site = "s1"
                wavelength = "w1"
            elif htd["sites"] == 1:
                plateName, well, wavelength = match.groups()
                site = "s1"
            elif htd["wavelength"]["number"] == 1:
                plateName, well, site = match.groups()
                wavelength = "w1"
            else:
                plateName, well, site, wavelength = match.groups()

            # seperate the directories in the path from the ending file
            dirpath = os.path.dirname(filepath)

            # remove the root part of the dirname
            dirpath = dirpath.replace(root, "", 1)

            # split dirpath into list of directories
            dirs = []
            while 1:
                head, tail = os.path.split(dirpath)
                dirs.insert(0, tail)
                dirpath = head
                if (head == "") or (head == os.path.sep): # head=="" is first level file, head==os.path.sep is second or greater level file
                    break

            dirs.append(file)

            # get the OME_Image
            new_filename = "_".join(file.split("_")[:-1])

            # truncate file name if too long
            if len(new_filename) > MAX_NAME:
                new_filename = truncate_name(dirs)

            # add row to the DataFrame
            #TODO: find out image name and MMA File path
            row = pd.DataFrame([[well,new_filename,file,filepath,json,site,wavelength,zStep,timepoint]],columns=columns)
            df = df._append(row, ignore_index=True)
        
    df.sort_values(by="File Name", inplace=True)
    return df

# NOTE: : doesn't close file after writing
def write_excel(file, df, sheet_number=2, cell="A14"):
    """Write DataFrame to an existing excel file on a specific sheet starting at a specific cell.

     Args:
        file: string file name.
        df: DataFrame table to write.
        sheet_number: integer zero-indexed index number of the sheet to write on.
        cell: cell to write DataFrame to.
     Returns:
        None
     Raises:
    """
    wb = xw.Book(file)
    sheet = wb.sheets[sheet_number]

    # Clear cells starting from the specified cell to the end of the used range
    clearStart = sheet.range(cell).address
    usedRange = sheet.used_range

    # define the range to clear
    clearRange = sheet.range(clearStart,(usedRange.last_cell.row,usedRange.last_cell.column))
    clearRange.clear()


    sheet[cell].options(index=False, header=False).value = df
    wb.save(file)

# NOTE: Using sheet index not sheet name because the unicode character U+0399 capital Greek Iota is present instead of U+0049 capital Roman I in some sheet names
# NOTE: : doesn't close file after reading
def read_excel(file, dataset_sheet=1, image_list_sheet=2, dataset_cell="C10", image_list_cell="B10"):
    """Read an excel file and extract dataset folder name and image file extensions from specific cells.

     Args:
        file: excel file.
        dataset_sheet: integer zero-indexed index number of the sheet containing dataset_cell.
        image_list_sheet: integer zero-indexed index number of the sheet containing image_list_cell.
        dataset_cell: string cell id of cell containing name of dataset flder.
        image_list_cell: string cell id of cell containing list of image file extensions.
     Returns:
        A tuple of (string dataset folder name, list of string image file extensions).
     Raises:
    """
    wb = xw.Book(file)
    sheet = wb.sheets[dataset_sheet]
    cell = sheet[dataset_cell]
    dataset = cell.value.strip()
    sheet = wb.sheets[image_list_sheet]
    cell = sheet[image_list_cell]
    extensions = cell.value
    extensions = extensions.split(" ")
    return (dataset, extensions)

def main(excel,isSPW):
    """Main method for running the script.

    Writes image file data like file and path names to the excel file with the name provided in the argument.

     Args:
        excel: string name of excel file script is being run from
     Returns:
        None
     Raises:
    """
    cwd = os.getcwd()

    dataset, extensions = read_excel(excel)
    if isSPW:
        #get htd file
        htd = getHtdFile(dataset)
        if htd:
            #using the htd file, get dictionary of valid and refused images. Retreive incomplete wells too
            validImages,rejectedImages = getImages(dataset,htd)
            incomplete = getIncompleteWells(htd, validImages)

            #TODO add error if there is something in rejectedImages or incomplete.
            if rejectedImages or incomplete:
                print("error")

            #get list of image names to be displayed on excel file
            validImageNames = getValidImageNames(validImages)
            df = create_DataFrame_SPW(os.path.join(cwd, dataset),validImageNames,extensions,htd)

        #if the htd file is null, process the images as if they were Project/Dataset    
        else:
            print("no htd found")
            df = create_DataFrame_ProjDataset(os.path.join(cwd, dataset), extensions)

    else:
        df = create_DataFrame_ProjDataset(os.path.join(cwd, dataset), extensions)

    write_excel(excel, df)

if __name__ == "__main__":
    #isSPW is used to determine if our images are in the SPW format or not
    isSPW = True
    #arg = sys.argv[1]
    main("Pazour_OMERO_import_template_wMacros_v06.xlsm",isSPW)
