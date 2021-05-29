from datetime import datetime

import pandas as pd


class CylinderRecords:
    def __init__(
        self,
        filepath: str,
        cylinder_name_column: str = "Tüp Seri No",
        date_col_name: str = "Tarih",
    ) -> None:
        self.cylinders = {}
        self.date_col_name = date_col_name
        self.cylinder_name_column = cylinder_name_column
        self.filepath = filepath

        print("reading xlsx file...")

        self.data = pd.read_excel(
            filepath,
            dtype=str,
            header=0,
            index_col=None,
            engine="openpyxl",
            usecols="A:J",
            na_filter=False,
        )
        print("reading file done")

    @staticmethod
    def update_cylinder_info(
        cylinders: dict, cylinder_name: str, timestamp: datetime, index: int
    ):
        if cylinder_name in cylinders:
            if cylinders[cylinder_name]["timestamp"] < timestamp:
                cylinders[cylinder_name] = {"timestamp": timestamp, "index": index}
        else:
            cylinders[cylinder_name] = {"timestamp": timestamp, "index": index}

    def filter_cylinder_records(self, out_file: str):
        print("filtering cylinder records...")
        cylinders = {}
        for index, row in self.data.iterrows():
            if index == 0:
                continue
            if index % 10000 == 0:
                print(f"first {index} rows done...")

            cylinder_name = row[self.cylinder_name_column]
            date_str = row[self.date_col_name]
            self.update_cylinder_info(cylinders, cylinder_name, date_str, index)

        indexes = [cylinders[cylinder_name]["index"] for cylinder_name in cylinders]
        filtered_data = self.data.iloc[indexes]
        print("filtered data type", type(filtered_data))
        print(f"writing to file {out_file} ...")
        filtered_data.to_excel(out_file, index=False)
        print("writing to file done.")


if __name__ == "__main__":

    # =========================================================
    # Set the following
    file_path = "/Users/umutkucukaslan/Downloads/tup3.xlsx"
    output_file_path = "/Users/umutkucukaslan/Downloads/tup3_out.xlsx"
    date_column_name = "Tarih"
    cylinder_name_column = "Tüp Seri No"
    # =========================================================

    records = CylinderRecords(file_path)
    records.filter_cylinder_records(out_file=output_file_path)
