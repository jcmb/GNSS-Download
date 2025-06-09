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

keyRemap= {
    "Time" : "GPS Week Seconds",
    "Week" : "GPS Week Number",
    "Type" : "fixType",
    "Mode" : "fix",
    "PDOP" : "PDOP",
    "Corr Age" : "Corr Age",
    "Used"  : "satellites",
    "Track" : "tracked",
    "East"  : "East Sigma",
    "North" : "North Sigma",
    "Hgt"   : "Up Sigma",
    "Velocity" : "Velocity",
    "Track Angle" : "Heading"
    }


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
    parser.add_argument("--ARP", "-A", help="ARP_Offset to be applied", type=float, default=0.0)
    parser.add_argument("--Trajectory", help="Output in trajectory format", type=float, default=0.0)
    parser.add_argument("--Verbose", "-V", help="Verbose", action="store_true")
    parser.add_argument("--Tell", "-T", help="Show Settings", action="store_true")

    parser = parser.parse_args()

    if parser.Tell:
        sys.stderr.write("KMZ:        {}\n".format(parser.KMZ_File))
        sys.stderr.write("Trajectory: {}\n".format(parser.Trajectory))
        sys.stderr.write("ARP:        {}\n".format(parser.ARP))
        sys.stderr.write("Save:       {}\n".format(parser.Save))
        sys.stderr.write("Verbose:    {}\n".format(parser.Verbose))

    if not os.path.isfile(parser.KMZ_File):
        print(
            f"Error: The file '{parser.KMZ_File}' does not exist or is not a valid file."
        )
        sys.exit(1)

    return vars(parser)


def parse_table(html):
    global keyRemap
    soup = BeautifulSoup(html, "html.parser")

    # Find all rows
    rows = soup.find_all("tr")

    # Extract required data
    data = {}
    for row in rows:
        cells = row.find_all("td")
#        print(cells[0].text.strip())
#        try:
#            print(cells[1].text.strip())
#        except:
#            print("")

        if len(cells) == 2:  # Only process rows with two cells
            key = cells[0].text.strip()
            value = cells[1].text.strip()
#            print(key)
#            pprint(cells)
#            print(value)
            if key in [
                "Time",
                "Week",
                "Type",
                "Mode",
                "PDOP",
                "Corr Age",
                "Used",
                "Track",
                "East",
                "North",
                "Hgt",
                "Velocity",
                "Track Angle"

            ]:
                remapped=keyRemap[key]
#                print(key,",", remapped)
                if not remapped in data:  # Work around that Hgt is used twice
                    data[remapped] = value
                else:
                    data["Up Sigma"] = value
            elif key == "UTC" :
                date, time = value.split("T")
                time = time.rstrip("Z")  # Remove trailing 'Z'
                data["Date"]=date
                data["Time"]=time
#                print("DT", date,time)

    if "Hgt" in data:
        if data["Hgt"].endswith("m"):
            data["Hgt"] = data["Hgt"][:-1]  # Remove m

    if "Corr Age" in data:
        if data["Corr Age"].endswith("s"):
            data["Corr Age"] = data["Corr Age"][:-1]  # Remove m

    if "East Sigma" in data:
        if data["East Sigma"].endswith("m"):
            data["East Sigma"] = data["East Sigma"][:-1]  # Remove m

    if "North Sigma" in data:
        if data["North Sigma"].endswith("m"):
            data["North Sigma"] = data["North Sigma"][:-1]  # Remove m

    if "Up Sigma" in data:
        if data["Up Sigma"].endswith("m"):
            data["Up Sigma"] = data["Up Sigma"][:-1]  # Remove m

    if "Heading" in data:
        if data["Heading"].endswith("°"):
            data["Heading"] = data["Heading"][:-1]  # Remove °

    if "Velocity" in data:
        if data["Velocity"].endswith("km/h"):
            data["Velocity"] = data["Velocity"][:-4]  # Remove m

    if "GPS Week Seconds" in data:
        if data["GPS Week Seconds"].endswith(" secs"):
            data["GPS Week Seconds"] = data["GPS Week Seconds"][:-5]  # Remove " secs"
#    pprint(data)
    return data


def parse_kmz(kmz_file, save_KML:bool, ARP_Offset:float=400.0, useKMZ:bool=True, trajectory:bool=False) -> None:

    if not isinstance(save_KML, bool):
        raise TypeError(f"save_KML must be a bool, but got {type(save_KML).__name__}")

    if not (isinstance(ARP_Offset, float)):
            raise TypeError(f"ARP_Offset must be a float or int, but got {type(ARP_Offset).__name__}")

    if not isinstance(useKMZ, bool):
        raise TypeError(f"save_KML must be a bool, but got {type(useKMZ).__name__}")

    if not isinstance(trajectory, bool):
        raise TypeError(f"trajectory must be a bool, but got {type(trajectory).__name__}")

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

            trajectory_fieldnames = [
                "GPS Week Seconds",
                "Latitude",
                "Longitude",
                "Height",
                "fix"]

            fieldnames = [
                "Date",
                "Time",
                "GPS Week Number",
                "GPS Week Seconds",
                "UTC Offset",
                "fix",
                "fixType",
                "fixInfo",
                "satellites",
                "tracked",
                "HDOP",
                "VDOP",
                "PDOP",
                "Latitude",
                "Longitude",
                "Height",
                "MSL separation",
                "Velocity",
                "Heading",
                "East Sigma",
                "North Sigma",
                "Up Sigma",
                "Corr Age"
            ]
            if trajectory :
                writer = csv.DictWriter(csvfile, fieldnames=trajectory_fieldnames,extrasaction='ignore')
            else:
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



                point=None
                coordinates=None
                point = placemark.find("kml:Point", namespace)
                if point is not None:
                    coordinates = point.find("kml:coordinates", namespace)
                    if coordinates is not None:
                        coordinates = coordinates.text.strip()
                        coordinates = coordinates.split(',')
                        if len (coordinates) == 3 :
                            details["Latitude"] =coordinates[0]
                            details["Longitude"] =coordinates[1]
                            details["Height"] =float(coordinates[2]) - ARP_Offset


                writer.writerow(details)

    except zipfile.BadZipFile:
        print("The file is not a valid KMZ.")
    except Exception as e:
        print(f"Error: {str(e)}")


def main():
    args = get_args()
    parse_kmz(args["KMZ_File"], args["Save"], args["ARP"],useKMZ=args["Save"])


if __name__ == "__main__":
    main()
