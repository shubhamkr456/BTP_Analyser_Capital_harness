import os
import pandas as pd

notes = os.getcwd() + "\\Notes\\Notes.xlsx"
sh1 = os.getcwd() + "\\Notes\\sh1.csv"


def filter(df, p, v):
    df = df[df[p.strip()] == v.strip()]
    return df


# return Matching Files
def list_files(directory):
    # List all files in the directory
    files = os.listdir(directory)
    xlsx_file = [file for file in files if file.endswith(".xlsx")]
    return xlsx_file


def underScore(x):
    if "_" in x:
        return True
    else:
        return False


def Dash(x):
    if "-" in x:
        return True
    else:
        return False


def notesText(term):
    notesdf = pd.read_excel(notes, dtype=str)
    x = notesdf[notesdf["Note"] == term]["Description"].to_string(index=False)
    return x


def process_ref(ref):
    """Process the reference by splitting at '_' or '-'."""
    if underScore(ref):
        try:
            return ref.split("_")[0]
        except Exception as e:
            print(f"Error processing underscore in '{ref}': {e}")
            return ref
    elif Dash(ref):
        try:
            return ref.split("-")[0]
        except Exception as e:
            print(f"Error processing dash in '{ref}': {e}")
            return ref
    else:
        return ref


def mdf1loc(x):
    if "GD" in x:
        x = x.replace(" ", "")
        return x
    else:
        return x




def term(to_ref):
    if "SP" in to_ref:
        tterm = "51"
    elif "S0" in to_ref:
        tterm = "900E"
    elif "SH" in to_ref:
        tterm = "900E"
    return tterm


# --Note returner
