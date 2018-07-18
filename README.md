# HPQC Processor

Convert Excel spreadsheets (the template included in this repository) to a combined spreadsheet capable of being imported into [HPQC](https://en.wikipedia.org/wiki/HP_Quality_Center).
The output can also be loaded by [Testrail](http://www.gurock.com/testrail/) as well.

## Getting Started

This is a command-line only application

### Prerequisites

Make sure you have the following things on your system

* [Python 3](https://www.python.org/getit/)
* [OpenPyxl](https://pypi.org/project/openpyxl/)
* [PyYAML](https://pypi.org/project/pyaml/)

## Usage

Make sure all the test scripts you want combined are in one folder and follow this command structure. Output type should be "HPQC" always for now

```
python3 hpqcprocessor.py [test scripts location] [config location] [hpqc header location] [output destination] [outputfile.xlsx] [Output type]
```

## Authors

* [**Clement Hathaway**](http://clement.nyc) - *Creator*
