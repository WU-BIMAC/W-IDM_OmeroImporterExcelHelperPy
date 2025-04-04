import sys
import os
import pandas as pd
import xlwings as xw

# maximum file name length
MAX_NAME = 255


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


def walk_files(root, extensions):
    """Navigates each file and folder beginning at the root folder and checks for image files.

    Args:
        root: path of the root project folder.
        extensions: list of string image file extensions.
    Yields:
        Yields file name and file path of each image file found.
    Raises:
    """
    for dirpath, dirnames, files in os.walk(root, topdown=True):
        for file in files:
            filename, extension = os.path.splitext(file)
            if extension in extensions:
                filepath = os.path.join(dirpath, file)
                # get accompanying json file
                json = None
                for f in os.listdir(dirpath):
                    name, ext = os.path.splitext(f)
                    if (ext == ".json") and (name.lower().count(file.lower()) > 0):
                        json = f
                yield (file, filepath, json)


def create_DataFrame(
    root,
    extensions,
    columns=["File Name", "New File Name", "File Path", "MMA File Path", "Tags"],
):
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
    walk = walk_files(root, extensions)

    for file, filepath, json in walk:
        # separate the directories in the path from the ending file
        dirpath = os.path.dirname(filepath)

        # remove the root part of the dirname
        dirpath = dirpath.replace(root, "", 1)

        # split dirpath into list of directories
        dirs = []
        while 1:
            head, tail = os.path.split(dirpath)
            dirs.insert(0, tail)
            dirpath = head
            if (head == "") or (
                head == os.path.sep
            ):  # head=="" is first level file, head==os.path.sep is second or greater level file
                break

        # separate tags with #
        tags = "#".join(dirs)

        dirs.append(file)

        # join directories with "_"
        new_filename = "_".join(dirs)
        if new_filename.startswith("_"):
            new_filename = new_filename.replace("_", "", 1)

        # truncate file name if too long
        if len(new_filename) > MAX_NAME:
            new_filename = truncate_name(dirs)

        # add row to the DataFrame
        # row = pd.DataFrame(
        #     [[file, new_filename, filepath, json, tags]], columns=columns
        # )
        df.loc[len(df)] = [file, new_filename, filepath, json, tags]
        # df = df.append(row, ignore_index=True)

    df.sort_values(by="File Name", inplace=True)
    return df


create_DataFrame("Dataset 1", [".czi"])


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
    sheet[cell].options(index=False, header=False).value = df
    wb.save(file)


# NOTE: Using sheet index not sheet name because the unicode character U+0399 capital Greek Iota is present instead of U+0049 capital Roman I in some sheet names
# NOTE: : doesn't close file after reading
def read_excel(
    file, dataset_sheet=1, image_list_sheet=2, dataset_cell="C10", image_list_cell="B10"
):
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
    extensions = cell.value.replace(" ", "")
    extensions = extensions.split(",")
    return (dataset, extensions)


def main(excel):
    """Main method for running the script.

    Writes image file data like file and path names to the excel file with the name provided in the argument.

     Args:
        excel: string name of excel file script is being run from
     Returns:
        None
     Raises:
    """
    # cwd = os.getcwd()
    directory = os.path.dirname(excel)

    dataset, extensions = read_excel(excel)
    df = create_DataFrame(os.path.join(directory, dataset), extensions)
    write_excel(excel, df)


if __name__ == "__main__":
    arg = sys.argv[1]
    main(arg)
