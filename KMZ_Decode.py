#!/usr/bin/env python3


import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
from pprint import pprint
import sys
import argparse
import os
import csv


from bs4 import BeautifulSoup


def get_args():

    parser = argparse.ArgumentParser(
        fromfile_prefix_chars="@", description="Trimble KMZ to CSV."
    )

    parser.add_argument("KMZ_File", help="KMZ file to convert to CSV", type=str)

    parser.add_argument(
        "--Save",
        "-S",
        help="Save the KML files from the KMZ. The CSV File will match the KML file",
        action="store_true",
    )
    parser.add_argument("--Verbose", "-V", help="Verbose", action="store_true")
    parser.add_argument("--Tell", "-T", help="Show Settings", action="store_true")

    parser = parser.parse_args()

    if parser.Tell:
        sys.stderr.write("KMZ :    {}\n".format(parser.KMZ_File))
        sys.stderr.write("Save:    {}\n".format(parser.Save))
        sys.stderr.write("Verbose: {}\n".format(parser.Verbose))

    if not os.path.isfile(parser.KMZ_File):
        print(
            f"Error: The file '{parser.KMZ_File}' does not exist or is not a valid file."
        )
        sys.exit(1)

    return vars(parser)


def parse_table(html):
    soup = BeautifulSoup(html, "html.parser")

    # Find all rows
    rows = soup.find_all("tr")

    # Extract required data
    data = {}
    for row in rows:
        cells = row.find_all("td")
        if len(cells) == 2:  # Only process rows with two cells
            key = cells[0].text.strip()
            value = cells[1].text.strip()
            if key in [
                "Week",
                "Time",
                "Mode",
                "Type",
                "Track",
                "Used",
                "PDOP",
                "Lat",
                "Lon",
                "Hgt",
            ]:
                if not key in data:  # Work around that Hgt is used twice
                    data[key] = value

    if "Hgt" in data:
        if data["Hgt"].endswith("m"):
            data["Hgt"] = data["Hgt"][:-1]  # Remove m

    if "Time" in data:
        if data["Time"].endswith(" secs"):
            data["Time"] = data["Time"][:-5]  # Remove " secs"
    return data


def parse_kmz(kmz_file, save_KML, useKMZ=True):
    # If Use KMZ is True then the CSV is based on the KMZ filename, which only works if there is a
    # Single KML file in the KMZ file. Which there is for Trimble KMZ files today.
    try:
        # Step 1: Open the KMZ file
        KMZ_dir = os.path.dirname(kmz_file)
        if KMZ_dir:
            KMZ_dir += os.sep

        with zipfile.ZipFile(kmz_file, "r") as z:
            # Step 2: Locate the KML file inside the KMZ
            kml_files = [name for name in z.namelist() if name.endswith(".kml")]

            if not kml_files:
                raise Exception("No KML file found in the KMZ.")

            # Extract the first KML file
            kml_file_name = kml_files[0]
            kml_data = z.read(kml_file_name)

            if save_KML:
                with open(KMZ_dir + kml_file_name, "wb") as file:
                    file.write(kml_data)
            if useKMZ:
                csvfile = open(kmz_file + ".csv", "w", newline="")
            else:
                csvfile = open(KMZ_dir + kml_file_name + ".csv", "w", newline="")
            fieldnames = [
                "Time",
                "Lat",
                "Lon",
                "Hgt",
                "Type",
                "HPrec",
                "VPrec",
                "Mode",
                "PDOP",
                "Tracked",
                "Used",
                "Week",
            ]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()

            # Step 3: Parse KML content
            tree = ET.parse(BytesIO(kml_data))
            #            ET.dump(tree)
            root = tree.getroot()

            # Step 4: Extract placemarks and coordinates
            namespace = {"kml": "http://www.opengis.net/kml/2.2"}
            placemarks = root.findall(".//kml:Placemark", namespace)

            for placemark in placemarks:
                description = placemark.find("kml:description", namespace)
                # Placemarks with co-ords have a description
                if description is None:
                    continue
                if isinstance(description, ET.Element):
                    details = parse_table(description.text)
                    writer.writerow(details)

    except zipfile.BadZipFile:
        print("The file is not a valid KMZ.")
    except Exception as e:
        print(f"Error: {str(e)}")


def main():
    args = get_args()
    parse_kmz(args["KMZ_File"], args["Save"], args["Save"])


if __name__ == "__main__":
    main()
