# HPQC Processor

Translate the multiple files of the template for use in making test scripts which are loadable by HPQC.
It also works with Testrail

## Getting Started

This is a command-line only application

### Prerequisites

Make sure you have the following things on your system

* Python 3
* OpenPyxl
* PyYAML

## Usage

Make sure all the test scripts you want combined are in one folder and follow this command structure. Output type should be "HPQC" always for now

```
python3 hpqcprocessor.py [test scripts location] [config location] [hpqc header location] [output destination] [outputfile.xlsx] [Output type]
```

## Authors

* [**Clement Hathaway**](http://clement.nyc) - *Creator*
