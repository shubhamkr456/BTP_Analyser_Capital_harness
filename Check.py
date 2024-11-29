import pandas as pd
import os
import numpy as np
import gc
import openpyxl
import xlsxwriter
from Functions import filter, list_files, underScore, Dash, process_ref, mdf1loc

pd.set_option("display.max_colwidth", None)


class DataProcessor:
    def __init__(self, loc_notes):
        self.mdf = None
        self.dnumber = None
        self.capital_file_path = None
        self.confi = None
        self.capital_file = None
        self.loc_notes = loc_notes

    def load_mdf(self, mdf_path, new_dnumber=None):
        # Only load mdf if new_dnumber is provided and different from the current dnumber
        if new_dnumber and new_dnumber != self.dnumber:
            self.dnumber = new_dnumber  # Update the dnumber
            self.mdf = pd.read_excel(mdf_path, dtype=str)
            self.mdf = self.mdf.fillna("")
        elif self.mdf is None:
            # If mdf is None (not loaded yet), load the data for the current dnumber
            self.mdf = pd.read_excel(mdf_path, dtype=str)
            self.mdf = self.mdf.fillna("")

    def load_capital_file(self, capital_file_path):
        self.capital_file = pd.read_excel(
            capital_file_path, dtype=str, header=3, sheet_name=1
        )
        self.capital_file = self.capital_file.fillna("")

    def clear_memory(self):
        # Clear capital_file and mdf if they are no longer needed
        self.capital_file = None
        self.confi = None

    def setConfi(self, confi):
        self.confi = confi

    def sdf_process(self, sdf):  # This removes SH1 & Sh2 wires
        COLUMNS = [
            "WIRE ID",
            "REFDES",
            "PIN",
            "REFDES.1",
            "PIN.1",
            "TERM PN",
            "TERM PN.1",
            "LTH",
            "NOTE CODE",
        ]
        sdf = sdf.loc[:, sdf.columns.intersection(COLUMNS)].copy()
        indices = sdf[
            ~(
                (sdf["WIRE ID"].str.contains("SH1"))
                | (sdf["WIRE ID"].str.contains("SH2"))
                | (
                    (sdf["WIRE ID"].str.contains("WW"))
                    & (sdf["WIRE ID"].str.contains("SH"))
                )
            )
        ].index
        sdf = sdf.loc[indices].copy()
        sdf.fillna("", inplace=True)
        return sdf

    def master_df(self):
        mdf1 = filter(self.mdf, "Config No", self.confi)
        COLUMNS = [
            "Bundle",
            "Wire No",
            "Gauge",
            "Color",
            "Length",
            "Item Refdes",
            "Pin",
            "Term Code",
            "Other-End Item",
            "Other-End Pin",
            "Other-End Term Code",
        ]
        mdf1 = mdf1.loc[:, COLUMNS]
        mdf1.fillna("", inplace=True)
        return mdf1

    def create_wire_id(self, mdf1):
        mdf1.fillna("", inplace=True)
        mdf1.loc[:, "WireId"] = np.where(
            mdf1["Wire No"].str.contains("WW"),
            mdf1["Wire No"].astype(str) + "-" + mdf1["Color"].astype(str),
            np.where(
                mdf1["Wire No"].str.contains("CR"),
                mdf1["Wire No"].astype(str) + "-" + mdf1["Gauge"].astype(str),
                np.where(
                    mdf1["Wire No"].str.contains("RR"),  # New condition
                    mdf1["Wire No"].astype(str),
                    "W"
                    + mdf1["Bundle"].astype(str)
                    + "-"
                    + mdf1["Wire No"].astype(str)
                    + "-"
                    + mdf1["Gauge"].astype(str)
                    + mdf1["Color"].astype(str),
                ),
            ),
        )
        return mdf1[
            [
                "WireId",
                "Length",
                "Item Refdes",
                "Pin",
                "Term Code",
                "Other-End Item",
                "Other-End Pin",
                "Other-End Term Code",
            ]
        ]

    def termfound(self, mdf1, wid, fterm):
        try:
            k = mdf1[mdf1["WireId"] == wid].index[0]
            if fterm == mdf1.loc[k, "Other-End Term Code"]:
                return mdf1.loc[k, "Term Code"]
            return mdf1.loc[k, "Other-End Term Code"]
        except:
            print(f"Error finding term for WireId: {wid}")
            return None

    def extract_pin_and_ref(self, ref):
        if "-" in ref:
            pin = ref.split("-")[1].strip()
            ref = ref.split("-")[0].strip()
        elif "_" in ref:
            pin = ref.split("_")[1].strip()
            ref = ref.split("_")[0].strip()
        else:
            pin = None  # Default case if neither "-" nor "_" is found
        return ref, pin

    def process_dataframes(self, df1, mdf1, wid, from_ref, to_ref, index):
        try:
            if not wid.startswith("W"):
                k = self.process_wire_id(wid, mdf1)
                # Case where it is different
            else:
                k = mdf1[mdf1["WireId"] == wid].index[0]
                
            from_ref, pin = self.extract_pin_and_ref(from_ref)  
            if not pin:  # If no pin found using delimiters, use df1
                pin = df1.at[index, "PIN"]
                
            to_ref, pin_1 = self.extract_pin_and_ref(to_ref)
            if not pin_1:  # If no pin found using delimiters, use df1
                pin_1 = df1.at[index, "PIN.1"]

            if from_ref == to_ref:
                print(f"From_ref:{from_ref}-Pin:{pin}; ToRef:{to_ref}:-{pin_1}")

                fterm, tterm = self.get_term_codes(
                    df1, mdf1, k, index, from_ref, to_ref, pin_1, pin
                )
            else:

                fterm, tterm = self.get_references(mdf1, k, from_ref, to_ref)
            return fterm, tterm
        except (IndexError, KeyError) as e:
            print(f"Error processing WIRE ID '{wid}': {e}")
            return "Update", "Update"

    def get_term_codes(self, df1, mdf1, k, index, from_ref, to_ref, pin_1, pin):
        convert_to_lowercase = lambda x: x[1:].lower() if x.startswith("*") else x
        mdf_pin = mdf1.loc[k, "Pin"]
        mdf_pin = convert_to_lowercase(mdf_pin)
        mdf_other_end_pin = mdf1.loc[k, "Other-End Pin"].strip()
        mdf_other_end_pin = convert_to_lowercase(mdf_other_end_pin).strip()
        # print(f"Final From Ref: '{from_ref}', To Ref: '{to_ref}', Pin: '{pin}', Pin_1: '{pin_1}'")

        # Case 1: Check if PIN.1 matches Other-End Pin
        if pin_1 == mdf_other_end_pin:
            print("Case 1")
            return mdf1.loc[k, "Term Code"], mdf1.loc[k, "Other-End Term Code"]

        # Case 2: Check if PIN matches Other-End Pin
        elif pin == mdf_other_end_pin:
            print("Case 2")
            return mdf1.loc[k, "Other-End Term Code"], mdf1.loc[k, "Term Code"]

        # Case 3: Check if PIN matches Pin
        elif pin == mdf_pin:
            print("Case 3")
            return mdf1.loc[k, "Term Code"], mdf1.loc[k, "Other-End Term Code"]
        # Case 4 when to has SP
        elif mdf_pin in from_ref:
            print("Case 4")
            return mdf1.loc[k, "Term Code"], mdf1.loc[k, "Other-End Term Code"]
        # case 5
        elif mdf_other_end_pin in from_ref:
            print("Case 5")
            return mdf1.loc[k, "Other-End Term Code"], mdf1.loc[k, "Term Code"]
        # case 6
        elif mdf_pin in to_ref:
            print("Case 6")
            return mdf1.loc[k, "Other-End Term Code"], mdf1.loc[k, "Term Code"]
        # case 7
        elif mdf_other_end_pin in to_ref:
            print("Case 6")
            return mdf1.loc[k, "Term Code"], mdf1.loc[k, "Other-End Term Code"]
        # Case 4: Fallback/default case (if no other matches)
        else:

            df1.at[index, "Manual Check"] = "Required"
            return mdf1.loc[k, "Term Code"], mdf1.loc[k, "Other-End Term Code"]

    def process_wire_id(self, wid, mdf1):  # returns index of master df
        try:
            # Handle "WW" in WireId
            if "WW" in wid:
                parts = wid.split("-", 2)
                if len(parts) == 3:
                    # Combine part[0] and part[1] with a hyphen
                    first_part = parts[0] + "-" + parts[1]
                elif len(parts) == 2:
                    # Use part[0] as the first part
                    first_part = parts[0]
                else:
                    print(f"Unexpected number of parts in WireId {wid}")

                if first_part:
                    matching_indices = mdf1[
                        mdf1["WireId"].str.contains(first_part)
                    ].index
                    if len(matching_indices) != 0:
                        return matching_indices[0]
                    else:
                        print(
                            f"No WireId containing '{first_part}' found in Master_df."
                        )
                        return None

            # Handle "CR" or "VR" in WireId
            elif "CR" in wid or "VR" in wid:
                first_part = wid.split("-", 2)
                first_part = first_part[0] + "-" + first_part[1]
                matching_indices = mdf1[mdf1["WireId"].str.contains(first_part)].index
                if len(matching_indices) > 0:
                    return matching_indices[0]
                else:
                    print(f"{wid} needs to be handled separately!")
                return None

            # Handle "RR" in WireId
            elif "RR" in wid:
                first_part = wid.split("-")[0]
                matching_indices = mdf1[mdf1["WireId"].str.contains(first_part)].index
                if len(matching_indices) > 0:
                    return matching_indices[0]
                return None

            # Handle other cases
            else:
                mdf2 = self.create_wire_id(self.mdf)
                if len(mdf2[mdf2["WireId"] == wid].index) == 0:
                    print(f"{wid} Not Found in Master Run_letter data!")
                    return None
                else:
                    return mdf2[mdf2["WireId"] == wid].index[0]

        except Exception as e:
            print(f"An error occurred: {str(e)}")
            return None

    def get_references(self, mdf1, k, from_ref, to_ref):
        if from_ref == mdf1loc(mdf1.loc[k, "Item Refdes"]):
            return mdf1.loc[k, "Term Code"], mdf1.loc[k, "Other-End Term Code"]
        elif from_ref == mdf1loc(mdf1.loc[k, "Other-End Item"]):
            return mdf1.loc[k, "Other-End Term Code"], mdf1.loc[k, "Term Code"]
        elif to_ref == mdf1loc(mdf1.loc[k, "Other-End Item"]):
            return mdf1.loc[k, "Term Code"], mdf1.loc[k, "Other-End Term Code"]
        elif to_ref == mdf1loc(mdf1.loc[k, "Item Refdes"]):
            return mdf1.loc[k, "Other-End Term Code"], mdf1.loc[k, "Term Code"]
        return "Update", "Update"

    def process_length(self, df1, mdf1, lth, wid, index, mdf):
        try:
            if not wid.startswith("W"):
                k = self.process_wire_id(wid, mdf1)
            else:
                k = mdf1[mdf1["WireId"] == wid].index[0]
            df1.at[index, "WDS LTH"] = mdf1.loc[k, "Length"]
            df1.at[index, "Length_Update"] = (
                "Same" if lth == mdf1.loc[k, "Length"] else "Different"
            )
        except:
            df1, mdf1 = self.handle_missing_wireid(df1, mdf1, wid, index, mdf)
        return df1, mdf1

    def handle_missing_wireid(self, df1, mdf1, wid, index, mdf):
        try:
            mdf_temp = self.create_wire_id(self.mdf)
            if not mdf_temp[mdf_temp["WireId"] == wid].empty:
                new_data = mdf_temp.loc[mdf_temp["WireId"] == wid]
                mdf1 = pd.concat([mdf1, new_data], ignore_index=True)
                df1.at[index, "WDS LTH"] = mdf_temp.loc[new_data.index[0], "Length"]
                df1.at[index, "Length_Update"] = (
                    "Same"
                    if df1.at[index, "LTH"] == df1.at[index, "WDS LTH"]
                    else "Different"
                )
            else:
                print(f"No matching WireId found in mdf_temp for {wid}")
        except Exception as e:
            print(f"Error processing mdf_temp for {wid}: {e}")
        return df1, mdf1

    def df1_apply(self, df1, mdf1, mdf):
        df1["note_add"] = ""
        df1["note_Text1"] = ""
        df1["note_add2"] = ""
        df1["note_Text2"] = ""
        df1["WDS LTH"] = ""
        df1["Length_Update"] = ""
        for index, row in df1.iterrows():
            wid = row["WIRE ID"]
            from_ref = row["REFDES"]  # Buggy ---->-----> Call it in  process dataframes
            # What is to be done.. Call in process_dataframe .. tap the pin into a variable and Update pin and pin_1 values.

            to_ref = row["REFDES.1"]  # buggy
            df1, mdf1 = self.process_length(df1, mdf1, row["LTH"], wid, index, mdf)
            fterm, tterm = self.process_dataframes(
                df1, mdf1, wid, from_ref, to_ref, index
            )
            df1.loc[index, "notes"] = f"{fterm},{tterm}"
            df1.loc[index, "note_add"] = f"note{fterm}@{row['REFDES']}:{row['PIN']}"
            df1.loc[index, "note_add2"] = (
                f"note{tterm}@{row['REFDES.1']}:{row['PIN.1']}"
            )
            df1.loc[index, "note_Text1"] = self.notesText(fterm)
            df1.loc[index, "note_Text2"] = self.notesText(tterm)

        return df1, mdf1

    def remove_extra_space(self, df):
        columns = ["Item Refdes", "Other-End Item"]
        for column in columns:
            df[column] = df[column].apply(
                lambda x: "".join(x.split()) if "GD" in x else x
            )
        return df

    def notesText(self, term):
        x = self.loc_notes[self.loc_notes["Note"] == term]["Description"].to_string(
            index=False
        )
        return x

    def split_arrange_notes(self, df):
        df["NOTE CODE"] = df["NOTE CODE"].str.replace(
            r"(\w+)-B", r"\1-F, \1-T", regex=True
        )
        df[["Cap_From_note", "Cap_To_note"]] = df["NOTE CODE"].apply(
            lambda x: pd.Series(self.sort_and_extract_notes(x))
        )
        return df

    # Function to sort and extract From_note and To_note
    def sort_and_extract_notes(self, notes):
        parts = notes.split(",")
        # Initialize variables
        from_note = ""
        to_note = ""

        for part in parts:
            if "-F" in part:
                from_note = part.replace("-F", "").strip()
            elif "-T" in part:
                to_note = part.replace("-T", "").strip()

        return from_note, to_note

    def mas_arrange_notes(self, df):
        df[["master_note_from", "master_note_to"]] = df["notes"].str.split(
            ",", expand=True
        )
        return df
