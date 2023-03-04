# @copyright Copyright (c) 2021-2023 Nicolas GRELLETY
# @license https://opensource.org/licenses/GPL-3.0 GNU GPLv3
# @link https://github.com/nillyr/octokonverter
# @since 1.0.0b

import argparse
import csv
import json
from pathlib import Path


def convert(input_file: Path, output_file: Path) -> None:
    """
    Converts the output file in csv format to a json file that can be imported by the add-in.
    """
    categories = []
    json_arr = []
    with open(input_file, "r") as csv_file:
        csv_reader = csv.DictReader(csv_file, delimiter=",", quotechar='"')
        for row in csv_reader:
            data = list(row.values())
            if not data[1] in categories:
                categories.append(data[1])
                json_arr.append(
                    {
                        "name_fr": data[1],  # the language doesn't really matter here
                        "name_en": data[1],
                        "rules": [
                            {
                                "reference": data[2],
                                "success": data[6],
                                "name_en": data[3],  # same here
                                "name_fr": data[3],
                                "level": data[4],
                            }
                        ],
                    }
                )
            else:
                for i in range(len(json_arr)):
                    for _, value in json_arr[i].items():
                        if value == data[1]:
                            json_arr[i]["rules"].append(
                                {
                                    "reference": data[2],
                                    "success": data[6],
                                    "name_en": data[3],
                                    "name_fr": data[3],
                                    "level": data[4],
                                }
                            )
                            break

    json_dict = {"categories": json_arr}
    with open(output_file, "w", encoding="utf-8") as json_file:
        json_file.write(json.dumps(json_dict, indent=4))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("-i", "--input", required=True, help="CSV input file")
    parser.add_argument("-o", "--output", required=True, help="JSON output file")
    args = parser.parse_args()

    convert(Path(args.input), Path(args.output))
