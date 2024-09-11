from fileinput import filename
from operator import index

import pandas as pd
from pathlib import Path

# merge all excel files from a directory
def mergefiles(output_file:str):
    # define a directory contain excel files
    directory = Path("file")
    # retrieve all excel files and convert the generator to list,
    # because During the first loop, the generator's content will be consumed
    files = list(directory.glob("*.xlsx"))
    print("Found files:")
    for f in files:
        print(f)
    dfs = []
    # Read each Excel file into a DataFrame and append it to the list
    for f in files:
        print(f"Reading {f}...")
        try:
            df = pd.read_excel(f)
            dfs.append(df)
        except Exception as e:
            print(f"Error reading {f}: {e}")
    if dfs:
        print("Concatenating DataFrames...")
        df = pd.concat(dfs, ignore_index=True)

        # write result to result Excel 文件
        print(f"Writing merged data to {output_file}...")
        df.to_excel(output_file, index=False)

        print(f"Data successfully written to {output_file}")
    else:
        print("No files were read. No output file was created.")

def merge_files_by_name(output_file: str, *files: str):
    print("Merge files ...")
    for file in files:
        print(file)
    dfs = []
    for file in files:
        try:
            df = pd.read_excel(file)
            dfs.append(df)
        except Exception as e:
            print("Error reading {file}:{e}")

    if dfs:
        print("Concatenating DataFrames...")
        df = pd.concat(dfs,ignore_index=True)
        print(f"Writing merged data to {output_file}...")
        df.to_excel(output_file,index=False)

        print(f"Data successfully written to {output_file}")
    else:
        print("No files were read. No output file was created.")


