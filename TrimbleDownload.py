#!/usr/bin/env python3

import os
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from pprint import pprint
from tqdm import tqdm
import argparse
from enum import Enum
import sys
from KMZ_Decode import parse_kmz


def get_args():

    parser = argparse.ArgumentParser(
        fromfile_prefix_chars="@", description="Trimble GNSS Download."
    )

    parser.add_argument("IP", help="GNSS receiver IP ")
    parser.add_argument(
        "Format",
        type=str.upper,
        choices=[fmt.short_code for fmt in GNSSFormat],
        help="Format of file to download",
    )

    # Define allowed RINEX versions
    allowed_versions = ["2.11", "2.12", "3.00", "3.02", "3.03", "3.04"]

    # Add argument with choices
    parser.add_argument(
        "--RINEX",
        choices=allowed_versions,
        default="3.04",
        required=False,
        help="Specify the RINEX version (allowed: "
        + ", ".join(allowed_versions)
        + " )",
    )

    parser.add_argument(
        "--Port",
        "-P",
        help="HTTP Port of the GNSS Receiver, Default: 80.",
        default=80,
        type=int,
    )
    parser.add_argument(
        "--Base",
        "-B",
        help="Base Directory to Download from, Default: Internal.",
        default="Internal",
    )
    parser.add_argument(
        "--Output", "-O", help="Directory to Download to.", default="./Downloads"
    )
    parser.add_argument(
        "--Max",
        "-M",
        help="Max number of downloads. Default Unlimited",
        default=None,
        type=int,
    )
    parser.add_argument(
        "--Recursive",
        "-R",
        help="Download from the Base directory and below",
        action="store_true",
    )
    parser.add_argument(
        "--Delete",
        "-D",
        help="Delete file after downloading, Not implemented",
        action="store_true",
    )
    parser.add_argument(
        "--Clobber", "-C", help="Overwrite Existing", action="store_true"
    )
    parser.add_argument(
        "--Quite", "-Q", help="Quite, do not show progress", action="store_true"
    )
    parser.add_argument(
        "--NoRename",
        help="Do not rename files to the name the browser would download.",
        action="store_true",
    )
    parser.add_argument("--DryRun", help="DryRun, do not delete.", action="store_true")
    parser.add_argument("--Verbose", "-V", help="Verbose", action="store_true")
    parser.add_argument("--ARP", "-A", help="ARP_Offset to be applied to CSV Files", type=float, default=0.0)
    parser.add_argument(
        "--Tell", "-T", help="Show Settings in use.", action="store_true"
    )

    parser = parser.parse_args()

    if parser.Tell:
        sys.stderr.write("IP:        {}\n".format(parser.IP))
        sys.stderr.write("Port:      {}\n".format(parser.Port))
        sys.stderr.write("Format:    {}\n".format(parser.Format))
        sys.stderr.write("Base:      {}\n".format(parser.Base))
        sys.stderr.write("Output:    {}\n".format(parser.Output))

        sys.stderr.write("RINEX V:   {}\n".format(parser.RINEX))
        sys.stderr.write("Max:       {}\n".format(parser.Max))
        sys.stderr.write("Recursive: {}\n".format(parser.Recursive))
        sys.stderr.write("Delete:    {}\n".format(parser.Delete))
        sys.stderr.write("Clobber:   {}\n".format(parser.Clobber))
        sys.stderr.write("NoRename:  {}\n".format(parser.NoRename))
        sys.stderr.write("Quite:     {}\n".format(parser.Quite))
        sys.stderr.write("DryRun:    {}\n".format(parser.DryRun))
        sys.stderr.write("Verbose:   {}\n".format(parser.Verbose))
        sys.stderr.write("\n")

    # http://172.27.0.42:83/xml/dynamic/fileManager.xml?deleteFiles=/Internal&f0=R2_R750___202409301600.T04&f1=R2_R750___202409301400.T04

    return vars(parser)


class GNSSFormat(Enum):
    HATANAKA = ("HATANAKA", "Hatanaka RINEX, Observations File.")
    HATANAKAZ = ("HATANAKAZ", "Hatanaka RINEX, All Data Zip File.")
    KML = ("KML", "Google Earth (line)")
    KMP = ("KMP", "Google Earth (line & points)")
    CSV = ("CSV", "CSV Generated from the Google Earth (line & points) file")
    RINEX = ("RINEX", "RINEX, Observations File.")
    RINEXZ = ("RINEXZ", "RINEX, All Data Zip File.")
    T0X = ("T0X", "Trimble T02 or T04 Format")

    @classmethod
    def from_string(cls, text):
        try:
            return cls[text]
        except KeyError:
            raise ValueError(
                f"{text} is not a valid GNSS format. Choose from {[e.name for e in cls]}"
            )

    def __init__(self, short_code, description):
        self.short_code = short_code
        self.description = description

    def __str__(self):
        #        return f"{self.short_code}"
        return f"{self.short_code} - {self.description}"


def filepathFrom_content_disposition(
    download_url, download_dir, default_filepath, verbose=False, NoRename=False
):

    if NoRename:
        #        print("Skipping Getting file name from the server")
        return default_filepath
    #    if verbose:
    #        print("Getting file name from the server")

    try:
        response = requests.head(download_url)
    except:
        raise SystemExit("ERROR: Ctrl-c while getting file name from server\n")

    response.raise_for_status()  # Raise an error for bad responses (4xx, 5xx)

    # Get filename from Content-Disposition header
    content_disposition = response.headers.get("Content-Disposition")

    if content_disposition:
        # Extract filename from header
        #        print(content_disposition)
        filepath = (
            download_dir
            + os.sep
            + content_disposition.split("filename=")[-1].strip(' "')
        )
        if verbose:
            print(
                "Using filename from server: ",
                content_disposition.split("filename=")[-1].strip(' "'),
                file=sys.stderr,
            )
    else:
        filepath = default_filepath
        if verbose:
            print(
                "Using default filename: ",
                download_url,
                " ",
                default_filepath,
                file=sys.stderr,
            )
    return filepath


def download_file(
    server,
    url,
    download_dir,
    outputFormat,
    RINEX=None,
    verbose=False,
    skip=False,
    progress=False,
    NoRename=False,
):
    """Download a file from the URL to the specified directory."""
    filepath = os.path.join(download_dir, url.split("/")[-1])
    download_url = server + url

    if outputFormat == GNSSFormat.RINEX:
        download_url += f"?format=RNX&Ver={RINEX}"
        filepath = filepathFrom_content_disposition(
            download_url,
            download_dir,
            filepath[:-3] + f"RNX.{RINEX}.obs",
            verbose=verbose,
            NoRename=NoRename,
        )
    elif outputFormat == GNSSFormat.RINEXZ:
        if RINEX in ["3.03", "3.04"]:
            download_url += f"?format=Zipped-RNX-MIX&Ver={RINEX}"
        else:
            download_url += f"?format=Zipped-RNX&Ver={RINEX}"
        filepath = filepathFrom_content_disposition(
            download_url,
            download_dir,
            filepath[:-3] + f"RNX.{RINEX}.zip",
            verbose=verbose,
            NoRename=NoRename,
        )
    elif outputFormat == GNSSFormat.HATANAKA:
        download_url += f"?format=RNX-COMP&Ver={RINEX}"
        filepath = filepathFrom_content_disposition(
            download_url,
            download_dir,
            filepath[:-3] + f"HATANAKA.{RINEX}.obs",
            verbose=verbose,
            NoRename=NoRename,
        )
    elif outputFormat == GNSSFormat.HATANAKAZ:
        download_url += f"?format=Zipped-RNX-COMP&Ver={RINEX}"
        filepath = filepathFrom_content_disposition(
            download_url,
            download_dir,
            filepath[:-3] + f"HATANAKA.{RINEX}.zip",
            verbose=verbose,
            NoRename=NoRename,
        )
    elif outputFormat == GNSSFormat.KML:
        download_url += "?format=KMZ-Lines"
        filepath = filepathFrom_content_disposition(
            download_url,
            download_dir,
            filepath[:-3] + "lines.kmz",
            verbose=verbose,
            NoRename=True,  #
        )
    elif outputFormat == GNSSFormat.KMP:
        download_url += "?format=KMZ-LinesPoints"
        filepath = filepathFrom_content_disposition(
            download_url,
            download_dir,
            filepath[:-3] + "kmz",
            verbose=verbose,
            NoRename=True,  # The server always returns the filename.kmz so skip the check
        )
    elif outputFormat == GNSSFormat.CSV:
        download_url += "?format=KMZ-LinesPoints"
        filepath = filepathFrom_content_disposition(
            download_url,
            download_dir,
            filepath[:-3] + "kmz",
            verbose=verbose,
            NoRename=True,
        )
        # When a
    else:
        pass  # T0x
    #        outputFormat == GNSSFormat.T0x:
    #        download_url+="?format=RNX&Ver=3.04"
    #        filepath = filepath[:-3] + "24O"

    if not (skip and os.path.isfile(filepath)):
        print(f"Downloading: {filepath} from {server}", file=sys.stderr)
    else:
        if verbose:
            print(
                f"Skipping existing file, {filepath}, would download from {download_url}",
                file=sys.stderr,
            )

    # Streaming, so we can iterate over the response.
    if not (skip and os.path.isfile(filepath)):
        response = requests.get(download_url, stream=True)
        response.raise_for_status()  # Raise an error for bad responses (4xx, 5xx)

        # Sizes in bytes.
        total_size = int(response.headers.get("content-length", 0))
        block_size = 10240

        if progress:
            try:
                with tqdm(
                    total=total_size, unit="B", desc=download_url, unit_scale=True
                ) as progress_bar:
                    with open(filepath, "wb") as file:
                        for data in response.iter_content(block_size):
                            progress_bar.update(len(data))
                            file.write(data)
                        pass
            except:
                print("Aborting: Removing partial download {}".format(filepath))
                os.remove(filepath)
                raise
        else:
            try:
                with open(filepath, "wb") as file:
                    for data in response.iter_content(block_size):
                        file.write(data)
            except:
                print("Aborting: Removing partial download {}".format(filepath))
                os.remove(filepath)
                raise

        if progress:
            if total_size != 0 and progress_bar.n != total_size:
                raise RuntimeError("Could not download file")
            else:
                print("")
    return filepath


def get_files_from_directory(server, directory_url, recursive=False, verbose=False):
    """Fetch the directory listing and extract file URLs."""

    if verbose:
        sys.stderr.write("Getting Page: {}\n".format(server + directory_url))
    try:
        response = requests.get(server + directory_url)
        response.raise_for_status()  # Raise an error for bad responses (4xx, 5xx)
    except:
        raise SystemExit(f"ERROR: Connect to {server} Failed\n")

    if response.status_code != 200:
        raise SystemExit(f"Failed to retrieve directory: {server}{directory_url}\n")

    # Parse the HTML content
    #    print(response.text)
    soup = BeautifulSoup(response.text, "html.parser")

    # Extract all links
    links = soup.find_all("a", href=True)

    # Filter links to include only .T02 or .T04 files
    file_urls = []
    for link in links:
        file_name = link["href"]
        if file_name.endswith(".T02") or file_name.endswith(".T04"):
            file_urls.append(urljoin(directory_url, file_name))
        elif (".T02?" in file_name) or (".T04?" in file_name):
            pass
            # Skip the computed files
        elif len(file_name) <= len(directory_url):
            pass
        else:
            if recursive:
                if verbose:
                    print(f"Going into {file_name}")
                #                pprint(file_urls)
                file_urls.extend(
                    get_files_from_directory(server, file_name, recursive, verbose)
                )
    #                pprint(file_urls)

    return file_urls


def is_stdout_redirected():
    return not sys.stdout.isatty()


def main():
    # URL of the directory (adjust with your server)
    args = get_args()
    verbose = args["Verbose"]
    progress = not args["Quite"]
    if is_stdout_redirected():
        progress = False
    outputFormat = GNSSFormat.from_string(args["Format"])
    #    pprint(args)

    server = "http://{}:{}".format(args["IP"], args["Port"])
    directory_url = "/download/{}".format(args["Base"])

    # Folder to store downloaded files
    download_dir = args["Output"]
    os.makedirs(download_dir, exist_ok=True)

    if verbose:
        print(f"Fetching files from {directory_url} to {download_dir}")

    # Get file URLs
    file_urls = get_files_from_directory(
        server, directory_url, args["Recursive"], verbose
    )

    numberDownloads = 0
    maxDownloads = args["Max"]

    if file_urls:
        print(f"Found {len(file_urls)} file(s) to download.\n")
        for file_url in file_urls:
            if maxDownloads and (numberDownloads >= maxDownloads):
                print("Max Downloads {} reached.".format(maxDownloads))
                break
            if verbose:
                print(f"Processing {file_url}")

            try:
                downloaded_file = download_file(
                    server,
                    file_url,
                    download_dir,
                    outputFormat,
                    args["RINEX"],
                    verbose=verbose,
                    skip=not args["Clobber"],
                    progress=progress,
                    NoRename=args["NoRename"],
                )
                #            print(downloaded_file)
            except:
                print("Downloading aborted")
                break
            if outputFormat == GNSSFormat.CSV:
                #                print (downloaded_file+".csv")
                #                print (os.path.isfile(downloaded_file+".csv"))
                if args["Clobber"] or (
                    not os.path.isfile(downloaded_file + ".csv")
                ):
                    if verbose:
                        print(
                            "Downloaded File: {}. Converting to CSV {}, ".format(
                                downloaded_file, downloaded_file + ".csv"
                            ),
                            end="",
                        )
                    parse_kmz(downloaded_file, False,args["ARP"])
                    if verbose:
                        print("Converted.")
                pass
            numberDownloads += 1
    else:
        print("No .T02 or .T04 files found.")


if __name__ == "__main__":
    main()
