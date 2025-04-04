import csv
import os

class CSVLoader:
    def __init__(self, filename):
        self.filename = filename
        self.base_path = os.path.dirname(os.path.abspath(__file__))
        self.file_path = os.path.join(self.base_path, filename)

    def load(self):
        """Load CSV data into a list of rows."""
        data = []
        encodings = ['utf-8', 'latin1', 'windows-1252']
        for encoding in encodings:
            try:
                with open(self.file_path, newline='', encoding=encoding) as csvfile:
                    reader = csv.reader(csvfile)
                    for row in reader:
                        if row:
                            data.append(row)
                return data
            except UnicodeDecodeError:
                continue
            except FileNotFoundError:
                print(f"Warning: {self.filename} not found!")
                return []
        return []

    def load_dict(self, key_column, value_column=None):
        """Load CSV into a dictionary with specified key and optional value column."""
        data = self.load()
        if not data:
            return {}
        headers = data[0]
        key_index = headers.index(key_column)
        if value_column:
            value_index = headers.index(value_column)
            return {row[key_index]: row[value_index] for row in data[1:]}
        return {row[key_index]: row for row in data[1:]}