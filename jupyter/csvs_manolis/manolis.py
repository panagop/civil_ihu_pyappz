import os
import csv

# ΦΑΚΕΛΟΣ ΜΕ ΤΑ CSV
folder_path = r"D:\mypython_projects\civil_ihu_pyappz\jupyter\csvs_manolis\Csv_External_Names"

for filename in os.listdir(folder_path):
    if filename.lower().endswith(".csv"):
        file_path = os.path.join(folder_path, filename)
        rows = []

        # Ανάγνωση CSV
        with open(file_path, newline="", encoding="utf-8") as f:
            reader = csv.reader(f)
            for row in reader:
                if row:  # αν δεν είναι κενή γραμμή
                    last = row[-1].strip()
                    if last == "ΣΥΝΑΦΕΣ":
                        row[-1] = "0"
                    elif last == "ΙΔΙΟ":
                        row[-1] = "1"
                rows.append(row)

        # Εγγραφή πίσω στο ίδιο αρχείο
        with open(file_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerows(rows)

print("Ολοκληρώθηκε η αντικατάσταση σε όλα τα CSV.")