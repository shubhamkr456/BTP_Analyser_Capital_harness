import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import gc
import warnings
from Check import DataProcessor
from Run_letter_object import WireCheckApp
import numpy as np

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


class MyApp:
    def __init__(self, root, loc_notes, xy):
        self.root = root
        self.root.title("BTP REPORT ANALYSER FILE WISE")
        self.root.geometry("400x255")

        self.flag = False
        # Button configuration
        button_config = {
            "padx": 10,
            "pady": 10,
            "bd": 1,
            "relief": tk.FLAT,
            "font": ("Consolas", 12, "bold"),
        }

        # Create buttons
        tk.Button(
            self.root,
            text="Select Master",
            bg="#000000",
            fg="#FFFFFF",
            command=self.load_master,
            **button_config,
        ).pack(fill=tk.X)

        tk.Button(
            self.root,
            text="Select Capital BTP",
            bg="#DD0000",
            fg="#FFFFFF",
            command=self.load_capital,
            **button_config,
        ).pack(fill=tk.X)

        tk.Button(
            self.root,
            text="Generate Report",
            border=5,
            fg="#11534d",
            bg="#FFCC00",  # cbf3f0
            command=self.generate_report,
            **button_config,
        ).pack(fill=tk.X)
        self.status_label = tk.Label(self.root, text="")

        self.status_label.pack(side="bottom", pady=15, ipady=10)
        self.footer_label = tk.Label(
            root, text="Crafted by Shubham", font=("Consolas", 12), fg="#c4ad68"
        )
        self.footer_label.pack(side="bottom", anchor="se", pady=10)

        self.master_file_path = None
        self.capital_file_path = None
        self.output_file_path = None
        self.loc_notes = pd.read_excel(loc_notes, dtype=str)
        self.xy = pd.read_excel(xy, dtype=str).drop(columns="Note_Value").fillna("")

    def extract_integer_from_notex(self, notex):
        # Check if 'WL' is in the string
        index = notex.find("WL")
        if index != -1:
            # Grab the next 2 characters after 'WL'
            next_two_chars = notex[index + 2 : index + 4]
            try:
                # Convert to integer
                result = int(next_two_chars)
                return result
            except ValueError:
                # Handle cases where conversion to integer fails
                print(f"Could not convert '{next_two_chars}' to integer.")
                return None
        else:
            return None

    def find_term(self, term, tdf):
        return (
            tdf[tdf["Term Code"] == term]["Part Number"].values[0]
            if not tdf[tdf["Term Code"] == term].empty
            else None
        )

    def add_single_quote(self, val):
        if isinstance(val, str) and val.startswith("="):
            return f"'{val}"

        return val

    def load_master(self):
        self.master_file_path = filedialog.askdirectory(title="Select Master Directory")

        if not self.master_file_path:
            messagebox.showerror("Error", "No Folder selected for Master data.")

    def load_capital(self):
        self.capital_file_path = filedialog.askopenfilename(
            title="Select Capital BTP File", filetypes=[("Excel files", "*.xlsx")]
        )
        if not self.capital_file_path:
            messagebox.showerror("Error", "No file selected for Capital BTP.")

    def generate_report(self):
        if not self.master_file_path or not self.capital_file_path:
            messagebox.showerror(
                "Error", "Please select both Master and Capital BTP files."
            )
            return

        # Load and process data
        dp = DataProcessor(loc_notes=self.loc_notes)
        # Drawing Number
        Drawing_number = self.capital_file_path.split("BTP")[1].split("-")
        dnumber = Drawing_number[0]
        Confi = Drawing_number[1].split("_")[0]
        # MAster file for Notes And Length
        mdf = self.master_file_path + "\\" + dnumber + ".xlsx"

        dp.load_mdf(mdf_path=mdf, new_dnumber=dnumber)
        dp.load_capital_file(self.capital_file_path)
        dp.setConfi(Confi)
        mdf1 = dp.master_df()
        mdf1 = dp.create_wire_id(mdf1)
        mdf1 = dp.remove_extra_space(mdf1)
        df1 = dp.sdf_process(dp.capital_file)
        mdf = dp.mdf

        df1, mdf1 = dp.df1_apply(df1, mdf1, mdf)
        sdf = dp.split_arrange_notes(df1)
        sdf = dp.mas_arrange_notes(sdf)

        sdf["from_term_part"] = sdf["master_note_from"].apply(
            lambda x: self.find_term(x, self.xy)
        )
        sdf["to_term_part"] = sdf["master_note_to"].apply(
            lambda x: self.find_term(x, self.xy)
        )

        sdf["From_note_Check"] = sdf["Cap_From_note"] == sdf["master_note_from"]
        sdf["To_note_Check"] = sdf["Cap_To_note"] == sdf["master_note_to"]
        sdf["Verdict"] = (sdf["TERM PN"] == sdf["from_term_part"]) & (
            sdf["to_term_part"] == sdf["TERM PN.1"]
        )

        # sdf update length wise.self. extract_integer_from_notex
        sdf["wl_value_add"] = sdf["note_add"].apply(self.extract_integer_from_notex)
        sdf["wl_value_add2"] = sdf["note_add2"].apply(self.extract_integer_from_notex)
        # Converting to numeric values
        sdf["LTH"] = pd.to_numeric(sdf["LTH"], errors="coerce")
        has_wl = sdf["note_add"].str.contains("WL", na=False) | sdf[
            "note_add2"
        ].str.contains("WL", na=False)

        match_condition = (sdf["wl_value_add"] == sdf["LTH"]) | (
            sdf["wl_value_add2"] == sdf["LTH"]
        )
        sdf.loc[has_wl & match_condition, "Length_Update"] = "Same"
        sdf.loc[has_wl & ~match_condition, "Length_Update"] = "Different"
        sdf["LTH"] = sdf["LTH"].astype(str)

        # Drop temporary columns
        #print(sdf.dtypes)
        sdf.loc[sdf["WIRE ID"].str.endswith("SH"), "Length_Update"] = "Same"
        #sdf.to_excel("Lth.xlsx")
        sdf = sdf.drop(columns=["wl_value_add", "wl_value_add2"])

        sdf_filtered = sdf.loc[
            ~(
                sdf["From_note_Check"]
                & sdf["To_note_Check"]
                & (sdf["Length_Update"] == "Same")
            )
        ]
        sdf_filtered.loc[sdf["To_note_Check"] == True, ["note_add2", "note_Text2"]] = ""
        sdf_filtered.loc[sdf["From_note_Check"] == True, ["note_add", "note_Text1"]] = (
            ""
        )

        if sdf_filtered.empty:
            sdf_filtered = sdf_filtered[
                [
                    "WIRE ID",
                    "LTH",
                    "WDS LTH",
                    "NOTE CODE",
                    "Length_Update",
                    "note_add",
                    "note_Text1",
                    "note_add2",
                    "note_Text2",
                    "master_note_from",
                    "master_note_to",
                ]
            ]

        else:
            if "Manual Check" not in sdf_filtered.columns:
                sdf_filtered["Manual Check"] = ""
            sdf_filtered = sdf_filtered[
                [
                    "WIRE ID",
                    "LTH",
                    "WDS LTH",
                    "NOTE CODE",
                    "Length_Update",
                    "note_add",
                    "note_Text1",
                    "note_add2",
                    "note_Text2",
                    "Manual Check",
                    "master_note_from",
                    "master_note_to",
                ]
            ]
        sdf_filtered = sdf_filtered.applymap(self.add_single_quote)

        Tcdf = sdf.loc[sdf["Verdict"] == False]
        Tcdf = Tcdf.rename(
            columns={
                "REFDES": "FROM REFDES",
                "REFDES.1": "TO REFDES",
                "TERM PN": "CAP FROM TERM PN",
                "TERM PN.1": "CAP TO TERM PN",
                "from_term_part": "Correct_from_term_part",
                "to_term_part": "Correct_To_Term_part",
            }
        )
        Tc_Columns = [
            "WIRE ID",
            "FROM REFDES",
            "PIN",
            "master_note_from",
            "CAP FROM TERM PN",
            "Correct_from_term_part",
            "TO REFDES",
            "PIN.1",
            "master_note_to",
            "CAP TO TERM PN",
            "Correct_To_Term_part",
            "Verdict",
        ]

        Tcdf = Tcdf[Tc_Columns]
        Tcdf = Tcdf.convert_dtypes()
        Tcdf.reset_index(drop=True)
        Tcdf = Tcdf.applymap(self.add_single_quote)
        Tcdf = Tcdf.fillna("")

        print(Tcdf.columns)

        if Tcdf.empty:
            print(f"Term Code Sheet Empty!")

        else:
            # Check this one
            Tcdf["From Check"] = np.where(
                (Tcdf["Correct_from_term_part"] == Tcdf["CAP FROM TERM PN"]),
                True,
                False,
            )
            Tcdf["TO Check"] = np.where(
                (Tcdf["Correct_To_Term_part"] == Tcdf["CAP TO TERM PN"]), True, False
            )
            #print(Tcdf)
            # Add the split
            mask_suffix = Tcdf["Correct_from_term_part"].str.contains(r"-X")
            # print(f"checking types: {Tcdf.dtypes}")

            # Extract the left part of `from_term_part` where suffix is present
            left_from_term_part = (
                Tcdf["Correct_from_term_part"]
                .where(mask_suffix)
                .str.split("-", n=1, expand=True)[0]
            )

            # Extract the left part of `FROM TERM PN-->`
            left_from_term_pn = Tcdf["CAP FROM TERM PN"].str.split(
                "-", n=1, expand=True
            )[0]

            # Update the `Verdict` column based on the comparison
            Tcdf.loc[mask_suffix & (left_from_term_part == left_from_term_pn)] = True

            Tcdf["Correct_To_Term_part"] = Tcdf["Correct_To_Term_part"].astype(str)
            # ----To Case
            mask_suffix = Tcdf["Correct_To_Term_part"].str.contains(r"-X")
            left_from_term_part = (
                Tcdf["Correct_To_Term_part"]
                .where(mask_suffix)
                .str.split("-", n=1, expand=True)[0]
            )
            Tcdf["CAP TO TERM PN"] = Tcdf["CAP TO TERM PN"].astype(str)
            # Extract the left part of `FROM TERM PN-->`
            left_from_term_pn = Tcdf["CAP TO TERM PN"].str.split("-", n=1, expand=True)[
                0
            ]
            Tcdf.loc[(mask_suffix) & (left_from_term_part == left_from_term_pn)] = True

            Tcdf = Tcdf.loc[~(Tcdf["TO Check"] & Tcdf["From Check"])]
            Tcdf.drop(columns=["Verdict"], inplace=True)
            Tcdf = Tcdf.rename(
                columns={
                    "master_note_from": "WDS_From_Note",
                    "master_note_to": "WDS_TO_Note",
                }
            )

        # For Run_letters--------------------
        capitalDf = pd.read_excel(
            self.capital_file_path, dtype=str, header=3, sheet_name="WIRE LIST"
        )
        runletter_d_number = capitalDf["HARNESS"].iloc[0]
        capitalDf = None
        Run_master = f"{self.master_file_path}\\{runletter_d_number}.xlsx"

        r = WireCheckApp(
            master_file_path=Run_master, capital_file_path=self.capital_file_path
        )
        Run_letterdf = r.generate_report()
        self.flag = (
            (Tcdf.shape[0] == 0)
            and (sdf_filtered.shape[0] == 0)
            and (
                not (
                    ("Master_Run_letter" in Run_letterdf.columns)
                    or (Run_letterdf["Run_letter_Status"].eq("").any())
                )
            )
        )

        output_file_path = (
            os.path.dirname(self.capital_file_path) + "\\Analysed_Output\\"
        )

        if not os.path.exists(output_file_path):
            os.makedirs(output_file_path)

        output_path = f"{output_file_path}{dnumber}-{Confi}_Analysed.xlsx"

        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            Tcdf.to_excel(writer, index=False, sheet_name="TermPart")
            workbook = writer.book
            worksheet1 = writer.sheets["TermPart"]
            header_format = workbook.add_format(
                {
                    "bold": True,
                    "text_wrap": True,
                    "valign": "top",
                    "bg_color": "#ff0400",
                    "font_color": "white",
                    "border": 2,
                    "align": "center",
                    "font_size": 12,
                }
            )

            worksheet1.set_column(0, len(Tcdf.columns) - 1, 20)
            for col_num, value in enumerate(Tcdf.columns.values):
                worksheet1.write(0, col_num, value, header_format)

            # -- sdf--Notes and Length-----------------------------------------------------------------------------
            sdf_filtered.to_excel(writer, index=False, sheet_name="Notes&Length")
            worksheet2 = writer.sheets["Notes&Length"]

            header_format = workbook.add_format(
                {
                    "valign": "vcenter",
                    "align": "center",
                    "bg_color": "#9900CC",
                    "bold": True,
                    "border": 2,
                    "font_size": 12,
                    "font_color": "#FFFFFF",
                }
            )
            textWrap = workbook.add_format({"text_wrap": "true", "border": 1})
            format = workbook.add_format(
                {
                    "font_color": "#000000",
                    "font_size": 11,
                    "border": 1,
                    "align": "center",
                }
            )
            worksheet2.set_column("A2:F1000", None, format)
            worksheet2.set_column("G:I", None, textWrap)
            worksheet2.set_row(0, 15)
            worksheet2.set_column("A:B", 21)
            worksheet2.set_column("B:F", 10)
            worksheet2.set_column("E:F", 16)
            worksheet2.set_column("F:G", 30)
            worksheet2.set_column("G:J", 26)
            worksheet2.set_column("H:J", 30)
            worksheet2.set_column("I:M", 20)
            for col_num, val in enumerate(sdf_filtered.columns.values):
                worksheet2.write(0, col_num, val, header_format)

            worksheet2.conditional_format(
                f"E2:E{len(sdf_filtered) + 1}",
                {
                    "type": "text",
                    "criteria": "containing",
                    "value": "Different",
                    "format": workbook.add_format(
                        {"bg_color": "#A8C77B", "bold": True, "font_color": "#FFFFFF"}
                    ),
                },
            )

            Run_letterdf.to_excel(writer, index=False, sheet_name="RunLetters")
            worksheet = writer.sheets["RunLetters"]
            header_format = workbook.add_format(
                {
                    "valign": "vcenter",
                    "align": "center",
                    "bg_color": "#2A3132",
                    "bold": True,
                    "border": 2,
                    "font_size": 12,
                    "font_color": "#FFFFFF",
                }
            )
            format = workbook.add_format(
                {
                    "font_color": "#000000",
                    "font_size": 11,
                    "border": 1,
                    "align": "center",
                }
            )
            worksheet.set_column("A2:F1000", None, format)
            worksheet.set_row(0, 15)
            worksheet.set_column("A:E", 21)

            for col_num, val in enumerate(Run_letterdf.columns.values):
                worksheet.write(0, col_num, val, header_format)

            worksheet.conditional_format(
                f"C2:C{len(Run_letterdf) + 1}",
                {
                    "type": "text",
                    "criteria": "containing",
                    "value": "Different",
                    "format": workbook.add_format(
                        {"bg_color": "#C7E0A2", "bold": True, "font_color": "#FFFFFF"}
                    ),
                },
            )

        messagebox.showinfo("Success", f"Report generated: {output_path}")
        if not self.flag:
            self.status_label.config(
                text="Changes required",
                fg="red",
                font=("Arial", 15, "bold"),
                bg="#000000",
            )  # 000000
        else:
            self.status_label.config(
                text="No Changes Required",
                fg="green",
                font=("Arial", 12, "bold"),
                bg="#BCFD49",
            )

        dp.clear_memory()
        gc.collect()


def main():
    root = tk.Tk()
    current_directory = os.getcwd()
    loc_notes_path = os.path.join(current_directory, "Notes", "Notes.xlsx")
    xy = os.path.join(current_directory, "Notes", "Term_TermPartnumber_C17.xlsx")
    app = MyApp(root, loc_notes=loc_notes_path, xy=xy)
    root.mainloop()


if __name__ == "__main__":
    main()
