import pandas as pd
import numpy as np
import xlsxwriter
import gc
import os


class WireCheckApp:
    def __init__(self, master_file_path="", capital_file_path=""):
        self.master_file_path = master_file_path
        self.capital_file_path = capital_file_path
        self.output_file_path = (
            os.path.dirname(self.capital_file_path) + "\\Analysed_Output\\"
        )
        self.mdf = pd.DataFrame()
        self.flag = False

    def generate_report(self):
        if not self.master_file_path or not self.capital_file_path:
            raise ValueError("Master and Capital BTP file paths must be provided.")

        try:
            return self.process_files()

        except Exception as e:
            print(f"Failed to generate report: {e}")

    def process_files(self):
        # Columns to filter
        master_columns = [
            "Bundle",
            "Wire No",
            "Gauge",
            "Color",
            "Config No",
            "Run Letters",
        ]

        # Load and process Capital BTP file
        capitalDf = pd.read_excel(
            self.capital_file_path, dtype=str, header=3, sheet_name="WIRE LIST"
        )
        capitalDf.fillna("", inplace=True)

        # Extract configuration details
        Drawing_number = self.capital_file_path.split("BTP")[1].split("-")
        dnumber = capitalDf["HARNESS"].iloc[0]
        Confi = Drawing_number[1].split("_")[0]

        # Load and filter Master data
        mdf1 = self.load_and_filter_data(
            self.master_file_path, dnumber, Confi, master_columns
        )
        mdf1 = self.create_wire_id(mdf1)

        # Process the Capital BTP data
        capitalDf = self.sdf_process(capitalDf)

        # Map and compare data
        for index, row in capitalDf.iterrows():
            wid = row["WIRE ID"]
            cap_sig = row["SIGNAL_CODE"]
            capitalDf = self.Signal_check(mdf1, wid, cap_sig, capitalDf, index)

        if ("Master_Run_letter" in capitalDf.columns) or (
            capitalDf["Run_letter_Status"].eq("").any()
        ):
            self.flag = True
        else:
            self.flag = False

        # Export to Excel
        # self.export_to_excel(capitalDf, dnumber, config=Confi)
        return capitalDf

        # Clear memory

    def load_and_filter_data(self, mdfloc, dnumber, Confi, master_columns):
        try:
            file_path = self.master_file_path
            self.mdf = pd.read_excel(file_path, dtype=str)
            mdf1 = self.mdf[self.mdf["Config No"] == Confi]
            mdf1 = mdf1[master_columns]
            return mdf1
        except FileNotFoundError:
            print("Master data file not found.")
            return pd.DataFrame(columns=master_columns)

    def missing_config_df(self):
        master_columns = [
            "Bundle",
            "Wire No",
            "Gauge",
            "Color",
            "Config No",
            "Run Letters",
        ]
        return self.create_wire_id(self.mdf[master_columns])

    def create_wire_id(self, mdf1):
        mdf1 = mdf1.copy()
        mdf1.fillna("", inplace=True)
        mdf1.loc[:, "WireId"] = np.where(
            mdf1["Wire No"].str.contains("WW"),
            mdf1["Wire No"].astype(str) + "-" + mdf1["Color"].astype(str),
            np.where(
                mdf1["Wire No"].str.contains("CR"),
                mdf1["Wire No"].astype(str) + "-" + mdf1["Gauge"].astype(str),
                np.where(
                    mdf1["Wire No"].str.contains("RR"),
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
        return mdf1[["WireId", "Run Letters"]]

    def sdf_process(self, sdf):
        COLUMNS = ["WIRE ID", "SIGNAL_CODE"]
        sdf = sdf.loc[:, COLUMNS]
        sdf = sdf[~sdf["WIRE ID"].str.contains("SH")]
        sdf.fillna("", inplace=True)
        return sdf

    def Signal_check(self, mdf1, wid, cap_sig, capitalDf, index):
        try:

            k = mdf1[mdf1["WireId"] == wid].index[0]
            run_letter = mdf1.loc[k, "Run Letters"]

            if run_letter == cap_sig:
                capitalDf.at[index, "Run_letter_Status"] = "Same"
            else:
                capitalDf.at[index, "Run_letter_Status"] = "Different"
                capitalDf.at[index, "Master_Run_letter"] = run_letter

        except IndexError:
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
                    return capitalDf
                # try Block
                try:
                    matching_indices = mdf1[
                        mdf1["WireId"].str.contains(first_part)
                    ].index
                    if len(matching_indices) == 1:
                        k = matching_indices[0]
                        run_letter = mdf1.loc[k, "Run Letters"]

                        if run_letter == cap_sig:
                            capitalDf.at[index, "Run_letter_Status"] = "Same"
                        else:
                            capitalDf.at[index, "Run_letter_Status"] = "Different"
                            capitalDf.at[index, "Master_Run_letter"] = run_letter
                    elif len(matching_indices) == 0:
                        print(
                            f"No WireId containing '{first_part}' found in Master_df."
                        )
                        pass  # No matching WireId found

                    elif mdf1.loc[matching_indices, "Run Letters"].nunique() == 1:
                        run_letter = mdf1.loc[matching_indices[0], "Run Letters"]
                        if run_letter == cap_sig:
                            capitalDf.at[index, "Run_letter_Status"] = "Same"
                        else:
                            capitalDf.at[index, "Run_letter_Status"] = "Different"
                            capitalDf.at[index, "Master_Run_letter"] = run_letter
                    else:
                        print(
                            f"Multiple WireIds containing '{first_part}' found in Master_df."
                        )

                except Exception as e:
                    print(f"An error occurred: {str(e)}")

            elif "CR" in wid or "VR" in wid:
                parts = wid.split("-", 2)
                first_part = parts[0] + "-" + parts[1]
                # trying to find index
                matching_indices = mdf1[mdf1["WireId"].str.contains(first_part)].index
                k = matching_indices[0]
                if len(matching_indices) > 0:
                    # Update the Run letter
                    run_letter = mdf1.loc[k, "Run Letters"]
                    if run_letter == cap_sig:
                        capitalDf.at[index, "Run_letter_Status"] = "Same"
                    else:
                        capitalDf.at[index, "Run_letter_Status"] = "Different"
                        capitalDf.at[index, "Master_Run_letter"] = run_letter
                else:
                    print(f"{wid} needs to be handled separately!")

            elif "RR" in wid:
                first_part = wid.split("-")[0]
                matching_indices = mdf1[mdf1["WireId"].str.contains(first_part)].index
                k = matching_indices[0]
                if len(matching_indices) != 0:
                    # Update the Run letter
                    run_letter = mdf1.loc[k, "Run Letters"]
                    if run_letter == cap_sig:
                        capitalDf.at[index, "Run_letter_Status"] = "Same"
                    else:
                        capitalDf.at[index, "Run_letter_Status"] = "Different"
                        capitalDf.at[index, "Master_Run_letter"] = run_letter
                else:
                    print(f"{wid} needs to be handled separately!")
            else:
                mdf2 = self.missing_config_df()
                # print(f"{wid} has Config_Mismatch")

                if len(mdf2[mdf2["WireId"] == wid].index) == 0:
                    print(f"{wid} Not Found in Master Run_letter data!")
                else:
                    k = mdf2[mdf2["WireId"] == wid].index[0]
                    run_letter = mdf2.loc[k, "Run Letters"]
                    if run_letter == cap_sig:
                        capitalDf.at[index, "Run_letter_Status"] = "Same"
                    else:
                        capitalDf.at[index, "Run_letter_Status"] = "Different"
                        capitalDf.at[index, "Master_Run_letter"] = run_letter
                # looking into Run Letters
        return capitalDf
